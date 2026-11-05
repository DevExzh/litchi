# 0301 — iWork structural leaf APIs (Numbers and Pages)

- Status: implemented on `uo2/iwork-structure`, cut from `feat/unified-ops-api`
  at `928b19b4db`, trunk merged at `cfef34af91` and `771a38c752`. Plan
  [0248](0248-coordinator-restart-plan.md) §6 Phase F (M6).
- Date: 2026-10-16
- Owner decisions applied: **D1**, **D2** (2026-10-15).
- Commits: `52e587279e` (litchi-numbers families), `c175e58e23`
  (litchi-iwa-archive directory publication, litchi-pages), `6922722ed0`
  (unified layer, tests, goldens), and this record.
- Did not touch `litchi-keynote` or `litchi-iwa`. In
  `crates/litchi-ops/src/format/iwork*` only `numbers.rs`, `pages.rs`, one
  `LeafError` variant in `error.rs` and the Pages arm of `Iwork::save_to` /
  `open_pages_directory` in `iwork.rs` changed.
- `performance_claim`: paired leaf/unified medians in §5 (release, one
  pinned core of a shared host, ABBA). No frozen-harness qualification.

Records 0297 and 0299 left the iWork structure as typed gaps. This record
closes the Numbers sheet and table lifecycle (requests 4, 5), the in-sheet
table order, Pages package-directory save-back (request 16) and the removal
of a Pages paragraph that anchors only body tables. It also fixes a writer
limitation those operations exposed: the Numbers table-cell writer refused
every multi-table workbook.

## 1. litchi-numbers (`52e587279e`)

**Shared machinery** (`package/structure.rs` and `structure/`). A structural
edit stages on a `Working` set. Each touched component is decompressed and
parsed once from the exact source, and its canonical framing is checked.
Objects are added, removed or rewritten in place. Untouched objects keep
their encoded bytes, and untouched members keep their compressed bytes.
`publish` does the rest in one pass:

- re-encodes the touched components;
- deletes any component the edit emptied;
- applies one registry transition to `Index/Metadata.iwa` through the shared
  `package_metadata_codec`: UUID removals and additions, external-reference
  removals and additions, the issued-identifier watermark, save tokens of
  the components whose registry records change, and component removals;
- drops the three stale root previews (all or none);
- reassembles once, reopens through the strict ingress, and verifies member
  locality.

Payload edits are wire-level (`structure/references.rs`): a repeated
`TSP.Reference` field is rewritten by reusing each retained record's bytes,
and header references follow through the core's exact transition
primitive. No generated-Prost production code was added.

The patch type keeps the exact source and target artifacts, so `inverse()`
restores the source byte for byte and `apply_x` accepts only that source.

**Families.** Each has the canonical `{Edit, Patch, Commit, Diagnostics,
Error, LimitKind}` set, and the glossary profiles register them (30
families).

| Family | Entry points | Operations |
|---|---|---|
| `sheet::lifecycle` | `edit_sheet_lifecycle`, `apply_sheet_lifecycle` | `add(name)` appends an empty sheet; `remove(selector)` |
| `table::lifecycle` | `edit_table_lifecycle(sheet)`, `apply_table_lifecycle` | `add(name, rows, columns)`; `remove(selector)` |
| `table::order` | `edit_table_order(sheet)`, `apply_table_order` | `move_table(selector, destination)` |

- **Sheet add.** Copies the first sheet's settings record without its private
  objects (drawables, deprecated header and footer storages, guide storage,
  headers, footers), sets the name, and appends the sheet to the
  document's sheet list and to the sidebar order as a new tree node. The
  leaf `sheet::order` requires that sidebar node, and the added sheet then
  moves and renames like any other.
- **Sheet remove.** Removes the sheet's tables first (table removal below),
  then detaches the sheet and its sidebar node. Its private closure goes
  with it: the objects whose every referrer is also removed, by the
  archive-header reference census in `structure/collect.rs`. A sheet with
  charts, media or shapes is refused (`SheetNotEmpty`). So is a workbook's
  last sheet (`LastSheet`).
- **Table remove.** Detaches the table drawable from its sheet and its
  sidebar node, and collects the private closure. Storage that a surviving
  table also references stays. The table's CalculationEngine owner family,
  their private dependency tiles, and every edge other owners hold into
  them go through `structure/formula.rs`, ported from the Pages body-table
  deletion owner (the same TSCE graph). A surviving formula that reads the
  table refuses the removal (`UnsupportedDependency`). Afterwards no
  surviving object may name a removed identifier (`verify_removed`).
- **Table add** (`table_lifecycle/add.rs`). The workbook's first table is
  the template, so styles, sizes and storage layout stay native:
  - the drawable and model are copied; the drawable is re-parented and,
    when it lands on the template's sheet, offset by 40 pt;
  - one emptied copy of each storage object the model owns: tiles with no
    rows, header buckets with no headers, data lists with no entries, an
    empty merge map, and a rebuilt UID map and stroke sidecar;
  - owners and providers beyond storage are dropped (sort, filter,
    category, pivot, hidden-state), and so are the title, caption, comment
    and the summary model;
  - fresh identity: a table UUID, registry UUIDs for clones of registered
    objects, and a formula owner registered in the dependency tracker and
    its owner map. The owner carries the native total and body range
    contexts, which the cache planner requires;
  - a sidebar node, and the external references the new objects need;
  - identities derive from the registry payload and the allocated
    identifiers (no clock, no randomness), so one edit of one source is
    byte-identical.
- **Table order.** Moves the table's record among the sheet's drawables:
  before the table that will follow it, or right after the last table.
  Other drawables keep their places. The sheet's sidebar table nodes
  follow when they list exactly the sheet's tables.

**Cell writer** (`table_cell_edit/{resolve,cache}.rs`). Two checks made
every multi-table workbook refuse cell writes. Record 0297 §4 saw one of
them as `InvalidSource` on the relocation fixture.

- The resolver required another table's owner to declare its `formula_owner`
  back-reference in its archive header. Native files never declare it. The
  check now runs only when a declaration is present.
- The cache planner required geometry for every table-backed owner, inert
  ones included. An owner with no formula in the final set, no dependency
  record, tile or range, and only the native range-context shape now needs
  no geometry. The unit test was split to cover that change: an inert
  foreign owner passes, and one holding a dependency tile is still refused.

A workbook with formulas in another table still refuses a write
(`UnsupportedDependency { FormulaCache }`): the planner needs every
formula-bearing table's geometry.

## 2. litchi-iwa-archive and litchi-pages (`c175e58e23`)

- `litchi_iwa_archive::publication::replace_directory_with` is the
  directory form of `replace_with`. It runs four steps:
  - the callback fills a fresh sibling staging directory;
  - every file and directory in it is synchronized;
  - an existing destination directory is renamed aside and the staging
    directory takes its name;
  - the old directory is then deleted.

  A failure before the swap leaves the destination untouched, and a failed
  swap renames the old directory back. The crate-boundary audit keeps
  `fs::rename` and synchronization inside this boundary, so the Pages leaf
  only writes files.
- `Package::save_directory(path)` writes `Index.zip`, which holds the
  `Index/` members with their records unchanged (one reassembly that
  deletes the other members), and every other member as a file. Opening
  the directory again gives back the exact single-file bytes.
  `SaveError::{Bundling, Bundle}` are new; a destination that exists but
  is not a directory is `SaveError::Publication`.
- `Package::section_paragraph_tables(section)` lists, per paragraph, the
  catalog positions of the body tables it anchors. It uses the table
  targets' anchor offsets and the section's body offset.

## 3. Unified layer (`6922722ed0`)

| Verb | Element | Leaf | Status |
|---|---|---|---|
| `add` | `sheet` under `/` (`name`, default `Sheet N`; any placement) | `sheet::lifecycle` (+ `sheet::order`) | supported |
| `remove` | `sheet` (with its tables) | `sheet::lifecycle` | supported |
| `add` | `table` under a sheet (`name`, `rowCount`, `columnCount`; default `Table N`, 10 × 5; placement among the sheet's tables) | `table::lifecycle` (+ `table::order`) | supported; a table-less workbook is a typed gap |
| `remove` | `table` | `table::lifecycle` | supported |
| `move` | `table` within its sheet (`Start`/`End`/`Position`/`Before`/`After`) | `table::order` | supported |
| `remove` | Pages `p` whose anchors are all body tables | `table::deletion` + `section::paragraphs` | supported |
| `save()` | Pages opened from a directory | `save_directory` | supported (the directory is the origin) |

Capability gaps removed:

- Numbers: add/remove sheet, add/remove table, in-sheet table move.
- Pages: the anchored-paragraph gap is narrowed to paragraphs that anchor
  charts, images or footnotes.

New: Numbers `add table` "a workbook without a table". Errors stay typed,
through three new `EditError` variants and one `LeafError::PagesSave`.

## 4. Tests and evidence

- `crates/litchi-numbers/tests/structure.rs` (8 tests) covers the created
  workbook and the whole 37-file Numbers corpus:
  - sheet add on every workbook, inverse byte-exact, remove gives back the
    original names;
  - 35 table removals, inverse byte-exact; 4 refused, typed;
  - 22 sheet removals; 15 refused (charts, media);
  - a table added to all 37 workbooks: 24 then take a value (the rest
    refuse, typed, for formulas in another table), and 36 remove again;
  - in-sheet moves;
  - refusals.
- The full `litchi-numbers` suite (1,057 tests) and `litchi-pages` suite
  (684) pass. The new Pages test `packages_save_back_as_directories_byte_exactly`
  covers directory round trip, replacement, conversion and a file
  destination.
- `litchi-ops`: `numbers_unified.rs` `sheets_and_tables_are_added_and_removed`
  and `corpus_tables_round_trip_through_add_and_remove` (> 30 round trips);
  the in-sheet move with inverse; `pages_unified.rs` directory `save()` in
  place, and an anchored paragraph removed with its table. The full suite
  passes (2,296 tests).
- Phase G checklist (`results/0289/crud-checklist.*`, re-blessed): Numbers
  has 45 claimed cells, all 45 passed. It generated the seeds `create() +
  add sheet` and `create() + add sheet + add table`, so sheet and table
  add, remove and move are now reached.
- Corpus fidelity (`results/0289/fidelity.*`, re-blessed) adds a Numbers
  `table add+remove` structural round trip: 38 value-exact, 1 one-way
  (`formula-envelope-native`: the inverse removal refuses a range tile with
  several from-to ranges), 1 refused. Byte equality is not expected: the
  forward step deletes the previews.
- CLI (`cli.rs`): a sheet and a table added, read, and the sheet removed.
  Python (`test_write_api.py`): the same through `litchi_py`.
- Goldens (capabilities and registry, Numbers and Pages) were blessed in
  both feature sets with identical output; the OfficeCLI delta table was
  regenerated.

## 5. Performance

`crates/litchi-ops/examples/numbers_unified_bench.rs` gained the scenarios
`sheet_add`, `sheet_remove`, `table_add`, `table_remove` and `table_move`,
all on the bench grid.

- Build: `cargo build --release -p litchi-ops --no-default-features
  --features numbers --example numbers_unified_bench`. Binary SHA-256
  `2a1913d98d12a52f248064221299d861f8796403b4a92cb63142ee5725da63c2`.
- Run: `LITCHI_NUMBERS_BENCH_SAMPLES=7 LITCHI_NUMBERS_BENCH_ONLY=…
  taskset -c 13`, twice, on a host at load 8–10. That is 7 ABBA rounds and
  14 samples per leg.
- The leaf and unified outputs were byte-identical in every scenario.
- Raw samples: [`results/0301/numbers-structure-bench.jsonl`](results/0301/numbers-structure-bench.jsonl)
  (5 KB).

| Scenario | Leaf median | Unified median | Ratio (run 1 / run 2) |
|---|---:|---:|---:|
| `set` (reference) | 358–362 µs | 359–363 µs | 0.993 / 1.013 |
| `sheet_add` | 245–250 µs | 244–254 µs | 1.034 / 0.978 |
| `sheet_remove` | 222–223 µs | 222–223 µs | 0.999 / 1.001 |
| `table_add` | 412–413 µs | 412–413 µs | 0.998 / 1.004 |
| `table_remove` | 228–233 µs | 229–232 µs | 0.984 / 1.017 |
| `table_move` | 227–228 µs | 227–231 µs | 0.994 / 1.015 |

Every structural operation is within the 10 % budget. The Pages anchored
removal composes the existing measured transactions. The directory save is
filesystem-bound and was not timed.

## 6. Gaps left (follow-ups)

1. **Numbers rows and columns (request 6).** The host's insert/delete
   closure is 9–12k lines: formula dependency shifting, tiles, UID maps,
   stroke layers, merges and headers. Not ported; it remains a typed gap.
2. **Formulas on `create()` (request 18), narrowed.** The blank's owner
   lacks the range contexts. With them added (tried, not committed), the
   formula-metadata planner then fails its exact execution budget one
   reference short (`References`, observed 9, maximum 8). Fix the planner's
   reference estimate, then regenerate the blank.
3. **Numbers cell writes with formulas in another table.** The cache planner
   needs a complete registry of formula-bearing tables.
4. **Numbers table add on a table-less workbook** (bootstrap from a theme
   preset, as the host does), and **table removal on a range tile with
   several from-to ranges**.
5. **Unchanged:** Numbers cell text styles (7), Numbers comments (19),
   cross-table formulas.
6. **Pages:** section add/remove (11); text runs and formatting (13); a
   Pages cell writer and table creation (sharing the Numbers TST writer
   through `litchi-numbers-wire`); removing paragraphs that anchor charts,
   images or footnotes; moving anchored paragraphs.
7. **Sheet removal with charts, media or shapes.** Needs a chart and media
   collection that also covers data members.
8. **Untested in the apps.** No native-app acceptance is claimed for the
   added sheets and tables. Their records follow the native shapes and the
   strict ingress reopens them.

## 7. Verification

Run on the final tree (`6922722ed0` plus trunk `771a38c752`). Every command
used `CARGO_BUILD_JOBS=6` and `RUST_TEST_THREADS=8`. All exit statuses
were checked.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers -p litchi-pages -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-iwa-archive -p litchi-pages --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers / pages / pages,numbers / keynote,numbers / pages,keynote] --all-targets -- -D warnings` | 0 (each) |
| `cargo test -p litchi-numbers --all-features` | 0 (1,057) |
| `cargo test -p litchi-pages --all-features` | 0 (684) |
| `cargo test -p litchi-iwa-archive --all-features` | 0 (200) |
| `cargo test -p litchi-ops --all-features` | 0 (2,296, after the re-bless) |
| `cargo test -p litchi --all-features` | 0 (260) |
| `cargo test -p litchi-cli --all-features` | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (276) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers -p litchi-pages -p litchi-iwa-archive -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 (it first flagged direct `fs::rename` and sync publication in the Pages leaf; moved into `litchi_iwa_archive::publication`) |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile numbers-minimal / numbers-full / pages / ops-iwork` | 0 (each; 30 Numbers families) |
| `python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, then the same with every feature but `encryption` | 0, 0 (identical) |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI` | 0 |

The OfficeCLI corpus has no iWork input, so it was not re-run. The target
directory reached 33 GB at peak and was pruned to 12–13 GB, with test
executables removed between suites.
