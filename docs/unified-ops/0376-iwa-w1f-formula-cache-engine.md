# 0376 — W1-F: the shared formula-dependency and formula-cache engine

- Status: landed on branch `uo2/iwa-w1f`, cut from trunk
  `feat/unified-ops-api` at `0fe32e12a9`; trunk `08e1011419` (W1-A
  `d836e4d693`) merged.
- Date: 2026-10-26.
- Plan: [0363](0363-iwa-monolith-exit-plan.md) (W1 table families), the
  formula graph of [0368](0368-iwa-w0b-codecs-formula-graph-table-vocabulary.md)/
  [0368a](0368a-iwa-w0b-regression-fixes.md), harness and ledger
  [0369](0369-iwa-exit-parity-harness-and-ledger.md), the shared table
  engine of [0370](0370-iwa-w1a-numbers-tables.md); owner decision D1 and
  the [macOS queue](0363-mac-queue.md).
- `performance_claim`: the Numbers cell-write routes are unchanged
  (−0.04%); the Keynote and Pages refresh routes run 69% and 97% below their
  host routes (callgrind, §4).

## 1. What landed

| Commit | Crate | What |
|---|---|---|
| `6cf8091364` | `litchi-numbers-wire`, `litchi-numbers` | engine move (§2); Numbers unchanged |
| `7b23e07486` | — | trunk merge |
| `f5a81fcbc9` | `litchi-numbers-wire` | `formula_cache::refresh`, the refresh for owners that stage payloads by object (§3) |
| `792136b2fb` | `litchi-keynote` | `slide::drawable::Edit::set_table_cell` refreshes dependent caches |
| `2836947de4` | `litchi-pages` | `table::values::Edit::commit` refreshes dependent caches |
| `174e634dfc` | `litchi-iwa` (tests), docs | parity cases, cost probe, 4 ledger rows, 2 macOS queue rows |
| this commit | docs | this record |

No crate edge changed (`check_crate_boundaries.py` 0). `litchi-numbers-wire`
gains `prost` as a dev-dependency only (the moved tests build their
fixtures with it).

## 2. The engine move

`litchi-numbers` `package/table_cell_edit/{cache, formula_metadata,
formula_list}.rs` (14.7k lines with their tests) move with `git mv` to
`litchi_numbers_wire::{formula_cache, formula_dependencies, formula_list}`:
the final-state dependency graph, dependent-cache planner and strict scalar
evaluator; the exact removal/insertion rewrite of the formula-owner graph;
and the formula-list entry plan. They read and write payload bytes and
object identifiers only (no generated protobuf value in any public
signature). `litchi-numbers` keeps its glue (`cache_commit`, `resolve`,
`tile`, the commit path) and imports the moved modules under their old
local names.

Conformance to 0351 on the moved surface:

- every entry point keeps its required limits (`CacheLimits`,
  `formula_dependencies::Limits`, `formula_list::Limits`); the glossary pins
  them (`required_signatures`: `plan_final_cache`,
  `plan_final_cache_with_authored`, `plan_final_cache_with_logical_graph`,
  `collect_formula_hosts`, `prepare_graph`, `formula_list::prepare`);
- `formula_cache::Failure` is now `Error`, and `LimitExceeded { observed,
  maximum }` is `Limit { kind, observed, maximum }` with a typed
  `formula_cache::LimitKind` (one per `CacheLimits` field);
- `formula_dependencies::Error::Limit { resource: &'static str }` is
  `Limit { kind: LimitKind }`;
- `formula_list::Resource` is now `LimitKind`, and its field is `kind`;
- the authored cache is a borrowed `formula_cache::SuppliedCache` instead of
  `litchi_numbers::formula::CachedValue`, which `litchi-numbers` converts
  at its boundary.

Numbers behaviour is unchanged: every `litchi-numbers` test (1,019) and the
50 moved tests (now in `litchi-numbers-wire`) pass unchanged, and so do the
Numbers parity cases. One observable detail changed: Numbers used to map the
dependency rewrite's limit names to its own `LimitKind` by substring. Six
output-size checks (`source candidate upper` and the five `encoded …`
checks) are guarded by `max_output_bytes` but were reported as `WireWork`.
They now report `OutputBytes`. No test depends on this.

## 3. Keynote and Pages formula support

`formula_cache::refresh::plan(Source { engine, owners, table }, writes,
object, &Limits) -> Plan` plans the refresh before any payload changes, from
the source state the `object` resolver returns:

1. A write that `table_write::cell_formula_state` reports as Inert needs no
   plan. Inert means no formula reads the cell and the cell holds none, the
   cell writer's admission before this record. Owners the planner declines
   elsewhere in the document then cannot refuse an unrelated write.
   `every_corpus_document_inserts_a_table_or_refuses_typed` (Pages, on
   `body-table-cross-reference-native.pages`) pins this case.
2. Otherwise the engine, every owner and its cell-record and range-precedent
   tiles are collected, and so is every table owner (kind 1) with the
   geometry of its `TableModel`. The edited table's formula cells and
   formula-list entries are collected too; a segmented list is refused. Every
   formula-owning table gets a values baseline, read from its BNC cells with
   the Numbers owner's semantics: text is aggregate-ignored and a formula
   precedent is unsupported. Then `plan_final_cache` runs with the
   `StrictEvaluator`, with the writes as the final-state overlay.
3. `TileCaches::apply` rewrites the owner's staged tile. It writes number and
   Boolean results through `BncCell::set_formula_cached_{number,boolean}`,
   as the host's `formula_cache.rs` does.

Typed refusals: overwriting or clearing a formula cell (the planner's
removal), a dependent formula the evaluator does not support (a direct
text reference, a non-scalar function, a volatile, spanning or
UUID-referenced dependency), a segmented formula list and a tile outside
the shared writer all return `Error::UnsupportedDependency`. Keynote maps
it to `Unsupported::TableFormulas` and Pages to `Refusal::Formulas`; both
variants' documentation now says what they mean. One deliberate difference
from the Numbers owner: range-precedent tiles are planned (the Numbers owner
refuses a selected owner that has them). All three Pages formula fixtures
put their `SUM(B2:B3)` table in such an owner, and the host refreshes it.

## 4. Parity

| Test | Fixture, operation | Tier | Result |
|---|---|---|---|
| `keynote_slide_table_formula_refresh_parity` | `slide-table-read-native.key`: B2 = 77; D2 = B2 + 1 | B | pass; the leaf changes D2's cache, and both outputs read the same cache |
| `keynote_slide_table_formula_refresh_clear_parity` | same, clear B2 | B | pass |
| `pages_body_table_formula_refresh_parity` | `body-table-cells-native.pages` table 1: B2 = 41; B4 = SUM(B2:B3) | B | pass, cache changed and equal |
| `pages_body_table_formula_refresh_clear_parity` | same, clear B3 | B | pass |
| `keynote_slide_table_cell_write_corpus_parity` | 3 fixtures | B | 47 writes compared (46 before); the listed formula refusal is removed |
| `pages_body_table_cell_write_corpus_parity` | 5 fixtures | B | the 96 listed refusals of W1-A's case now compare and pass; the listed-refusal arm is removed |

Where the routes still differ:

- **Both refuse.** Clearing or replacing a formula cell: the host fails with
  "expected to remove one protobuf field 11" and the leaf refuses
  `Formulas`. Writing text into Keynote B2, which D2 reads directly: the
  host says "non-scalar cached value" and the leaf refuses `TableFormulas`.
- **The host refuses, the leaf edits (6 writes).** Text into B2 or B3 of the
  `SUM(B2:B3)` table of `body-table-{cells,merges,read}-native.pages`. The
  leaf follows the Numbers owner, where SUM ignores text, and the host
  refuses. Output differs from the host, so the Pages queue row says so.

Ledger: `keynote/editor/slide_tables.rs` `set_slide_table_cell` and
`clear_slide_table_cell` → `litchi_keynote::slide::drawable::Edit::set_table_cell`;
`pages/editor/tables/semantic.rs` `set_table_cell` and `clear_table_cell` →
`litchi_pages::table::values::Edit::set`; tier B, queued. The macOS queue has
two W1-F rows, Keynote and Pages, because there is no native evidence for a
Litchi-written formula cache.

## 5. Performance

Callgrind, release with debug information (`lto = true`), CPU 3 pinned,
instruction counts (Ir).

| Route | Before / host | After / leaf | Δ |
|---|---:|---:|---:|
| Numbers cell writes, total of both probe routes, ABBA mean (before `0fe32e12a9`, after `6cf8091364`) | 818,991,697 | 818,678,638 | −0.04% |
| · every stored cell of 7 Numbers fixtures set to 7, ×2 | 577,437,470 | 577,159,995 | −0.05% |
| · `basic.numbers` formula lifecycle (author C3 = SUM(B3, 8), write B3, replace, clear), ×2 | 241,554,228 | 241,518,643 | −0.01% |
| Keynote B2 = 77 with D2 refresh, ×2: host `set_slide_table_cell` vs leaf `set_table_cell` | 636,323,174 | 195,062,966 | **−69%** |
| Pages B2 = 41 with B4 refresh, ×2: host `set_table_cell` vs leaf `edit_body_table_values` | 2,556,586,666 | 83,446,135 | **−97%** |

The leaf Keynote count also includes the probe's drawable-index search,
which runs the leaf route on drawables that refuse. The Numbers probe was a
scratch integration test of `litchi-iwa` built at both commits; it is not
committed. The other probes are committed:

```sh
CARGO_PROFILE_RELEASE_DEBUG=1 cargo test --release -p litchi-iwa --test parity --no-run
taskset -c 3 valgrind --tool=callgrind \
  '--toggle-collect=*formula_refresh_cases::host_refresh_keynote*' \
  target/release/deps/parity-<hash> --ignored --exact \
  formula_refresh_cases::formula_refresh_cost_abba --test-threads=1
# likewise leaf_refresh_keynote, host_refresh_pages, leaf_refresh_pages
```

Not measured: formula-free Keynote and Pages writes. They now also locate
the calculation-engine payload while collecting owners, but they skip the
planner after the same `cell_formula_state` check as before.

## 6. Gaps and follow-ups

| Gap | Owner |
|---|---|
| Formula-cell overwrite or clear (dependency removal through `formula_dependencies` and formula-list release) is still refused in Keynote and Pages, as in the Numbers scalar path; the host fails on the same fixtures, so no oracle exists | W1-F follow-up, together with the Numbers removal path |
| Formula authoring (writing a formula) in Keynote and Pages tables | follows the Numbers `formula_author` port |
| Segmented formula lists are refused | W1-F follow-up |
| Dependents in other tables are not refreshed: only the edited table's formula cells are planned, as in the Numbers owner, so a write a formula in another table reads is refused (`ExternalOwner`) | W1-F follow-up |
| The 6 text writes the host refuses (§4) and both new routes in general need the macOS run | macOS queue |
| `formula_cache::refresh` has no unit tests in `litchi-numbers-wire`; it is covered by the Keynote and Pages tests and the parity target | W1-F follow-up |
| Batch rows `set_slide_table_cells` and Pages `set_table_cells` stay `host-only` | W1-C, W1-B |

## 7. Verification

Build env of the coordinator rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`), on the branch with trunk `08e1011419` merged. Both
pre-commit size checks printed nothing for every commit.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-keynote -p litchi-pages -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features numbers`, `pages`, `keynote` | 0, 0, 0, 0 |
| `cargo test --all-features`: `-p litchi-numbers-wire` (263) / `-p litchi-numbers` (1,019) / `-p litchi-keynote` (1,184) / `-p litchi-pages` (716) / `-p litchi-iwa` (1,941; `parity` target 148 passed, 7 ignored probes) / `-p litchi` | 0 each |
| `cargo test -p litchi-ops --all-features`: `--lib`, `--doc`, nine groups of up to 20 `--test` targets | 0 each, after the two updates below |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers-wire -p litchi-numbers -p litchi-keynote -p litchi-pages` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/iwa_exit_ledger.py --check` | 0; 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger` (1,289 tests) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` | 0, 0, 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

Two litchi-ops tests changed with the behaviour, both intended:

- `pages_unified::tables_read_cells_and_remove` asserted that `set` of a cell
  in `body-table-cells-native.pages` table 1 is refused. That table takes part
  in a formula graph, and the write now succeeds and reads back.
- `crud_matrix::crud_checklist_per_format` was re-blessed with `LITCHI_BLESS=1`.
  One row of `results/0289/crud-checklist.tsv` changed: Pages `cell set` now
  succeeds on its first candidate fixture, `body-table-read-native.pages + add
  table`.

The capability and registry goldens and the OfficeCLI corpus pass unchanged,
so they were not re-blessed.
