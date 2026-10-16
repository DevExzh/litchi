# 0309 — iWork element gaps (rows and columns, character styling, table cells)

- Status: implemented on `uo2/iwork-elements`. Record 0306's branch was
  merged in at `5957672068`, and trunk `feat/unified-ops-api` at
  `cae9d95d34` and `972ff753ec` (records 0305, 0308; trunk head
  `1178eb235c`). Plan [0248](0248-coordinator-restart-plan.md) Phase F (M6).
- Date: 2026-10-16
- Owner decisions applied: **D1**, **D2** (2026-10-15): the iWork crates are
  edited directly.
- Commits: `faa65518c3`, `4d692176f6` (Numbers topology), `7069ed86b7`
  (unified rows and columns), `85878f25d4`, `56ad91675d`, `ef2765a729`,
  `81459222ab`, `d460a40d89`, `1ef74116c5` (character styling),
  `918eea9cc6`, `8e19a70696`, `195bfccc1b`, `3118f1b6ca` (table cells),
  `f3ade4042f` (Keynote media and placeholders), `16f1bd3c10` (request 18),
  `23cf8e176d`, `7cc3d04195`, `34a01b90ab` (tests, checklist, goldens),
  `4c611be0e5` (benches), and this record.
- `performance_claim`: paired leaf/unified medians in §6 (release build, one
  pinned core of a shared host, ABBA, two runs). No frozen-harness
  qualification.

The goal ranks the work by user value: Numbers rows and columns first, then
text styling in all three formats, then tables, then the rest. Every edit
below keeps untouched bytes exact. Every edit returns an exact-source patch
whose `inverse()` restores the source byte for byte. Content the leaf does
not port is refused with a typed error at the element.

## 1. Numbers rows and columns (`table::topology`, request 6)

`litchi_numbers::table::topology` provides `Package::{edit_table_topology,
apply_table_topology}` and `Edit::{insert_row, insert_column, remove_row,
remove_column}`. Positions are relative to the header, body and footer
sections (`RowInsertion::{header, body, footer}`, `ColumnInsertion::{header,
body}`, and the matching deletions). One edit stages several operations and
publishes them in one commit.

What it rewrites:

- tiles, both modern and pre-BNC: rows are spliced, cell offsets shifted,
  and missing rows materialised;
- the row and column info buffers and the header row and column counts;
- the table's layout cache and stroke layers;
- the `TST.TableDataList` references held by removed cells: each released
  key decrements its entry, and an entry that reaches zero is dropped along
  with the objects it referenced (`table_topology/lists.rs`).

`litchi_numbers_wire` gained `ListKind` and `BncCell::list_keys` so that a
removed cell's keys can be enumerated.

Refused, typed (`Error::Unsupported(Unsupported::…)`):

- `Formulas` and `ReferencedByOtherOwners`: the table holds formulas, or a
  formula or chart elsewhere reads it. No formula-coordinate shift is
  ported.
- `Merges`, `Filters`, `Categories`, `Pivot`.
- `SortRules`, for column edits only.
- `HiddenAxis`, `Borders`, `Comments`: removing a hidden, bordered or
  commented row or column.
- `StorageLayout`: a non-native tile layout.

Other refusals are `PositionOutOfRange`, `SectionFull`, `LastAxis` and
`CellClear`.

Corpus (`crates/litchi-numbers/tests/topology.rs`): 28 corpus tables insert
and remove a row and a column and round trip, 22 of them removing cells that
hold content. The rest refuse with the typed reasons above.

**Unified.** `row` and `col` (alias `column`) are elements under a Numbers
table, addressed as `row[N]` and `col[N]`. `add` takes `index`, `before` or
`after`, and `remove` takes the path. Since the 0305 merge, rows and columns
are slots of the lazy table part after the merged ranges. Their items are
built together on the first read of either one, so opening a table costs
two counts.

## 2. Character styling (requests 7, 10, 13)

**Shared wire** (`litchi_iwa_text_wire::char_style`): `CharacterProperties`
(bold, italic, underline, strikethrough, size, font, colour) with
inheritance through the parent chain, plus:

- `char_style_runs` and `styles_in_span`;
- `restyle_span`, which splits the character-style table at the span edges;
- `character_style_variation` and the stylesheet helpers.

**Keynote** (`slide::drawable::Edit::set_text_style`, request 10). Writes
character-style variations into the storage's stylesheet and registers
their UUIDs and external references. `Unsupported::TextStyling` now means
only "the style sets nothing". The text runs resolve the parent chain.
Stylesheets whose header cannot take the transition fall back to the
stylesheet member (`d460a40d89`). Corpus: 45 text boxes and shapes styled.

**Pages** (`section::text_style`, request 13): `Package::{edit_section_text_style,
apply_section_text_style, section_text_runs}`, with its own staging and a
metadata rewrite. Corpus: 31 sections styled.

**Numbers** (`cell::text_style`, request 7): `Package::{edit_table_cell_text_style,
apply_table_cell_text_style, table_cell_text_style}`. The properties go into
a `ParagraphStyleArchive` variation of the cell's style, plus a style-list
entry. Restyling releases the old key. An empty cell is refused
(`EmptyCell`). Corpus: 33 cells styled.

**Unified.**

- Keynote `run`: `bold`, `italic`, `underline`, `size`, `color` and
  `font.latin` are readable and writable.
- Numbers `cell` and Pages `p`: `font.bold`, `font.italic`,
  `font.underline`, `font.strike`, `font.size`, `font.name` and
  `font.color` are set-only. Pages has no run elements in its tree yet, so
  styling is per paragraph.

The `font.*` parser is shared (`format/iwork/write.rs`).

## 3. Table cells (Pages, Keynote)

`litchi_numbers_wire::table_write` is the TST cell writer shared by Pages
and Keynote tables:

- `tile_set_cell`, `header_adjust`, `list_intern` and `list_release`;
- `cell_formula_state` and `table_formula_state`;
- merge-owner (kind 5) ranges are ignored.

The leaf families on top of it:

- **Pages** `table::values`: `Package::{edit_body_table_values,
  apply_body_table_values}`. Corpus: 13 body tables written; the 4 holding
  formulas are refused (`Formulas`).
- **Keynote** `slide::drawable::Edit::set_table_cell(slide, drawable, row,
  column, TableCellValue)`. Refusals: `Unsupported::{TableFormulas,
  TableStorage}`.

**Unified.** `tc` takes `text` or `value` (text, number, boolean, or null to
clear). A formula cell, or a cell a formula or chart reads, is refused so
that no calculation cache goes stale.

## 4. Keynote media and placeholders; request 18

- `remove` of a `video` or `audio` goes through `slide::media::lifecycle`
  (`Patch::Media`).
- `remove` of a title or body placeholder hides it on its slide, and
  `hidden` = false shows it again (`slide::placeholder` visibility). The
  `hidden` property is set-only. Before this record that `remove` was
  refused, and the CLI and Python tests expected the refusal; they now
  expect the hide.
- Request 18: the blank Numbers asset (`create/blank.numbers`) now carries
  its table owner's extents, and the formula-metadata preflight counts the
  owners. Formulas therefore work on a created workbook, and the unified
  create test expects success.
- Unrelated fix: `limits::Profile::new` was gated on
  `internal-iwork-source` but is used by directory opening, so
  `litchi-numbers` failed to build with default features. The gate is gone.

## 5. Tests, checklist, fidelity, goldens

- Unified tests: `numbers_unified.rs` (rows and columns, formulas on a
  created workbook), `keynote_unified.rs` (run styles, slide-table cells,
  media removal, placeholder hide and show), `pages_unified.rs` (paragraph
  fonts, body-table cells).
- CLI (`crates/litchi-cli/tests/cli.rs`):
  - insert and remove rows and columns, including the typed refusal on a
    table with merged cells;
  - `font.bold` on a Numbers cell, a Pages body-table cell, and a Keynote
    text-box run;
  - placeholder hide and show.
- Python (`test_write_api.py::test_iwork_element_writes_round_trip`): the
  same operations through `litchi_py`.
- CRUD checklist (`results/0289/crud-checklist.*`). Numbers rows and columns
  are tree nodes, so `get`, `find`, `query` and `view` reach them. The
  checklist gains a Pages table without formulas
  (`body-table-name-focused.pages`) and a Keynote seed with a two-paragraph
  text box; for writes, drawables are tried before placeholders. Newly
  claimed cells, all `ok`:
  - Numbers `row` and `col`: get, find, query, view, add, remove;
  - Pages `cell` set;
  - Keynote `run` set and `tc` set;
  - Keynote `video`, `audio` and `placeholder` remove.

  Keynote `connector` and `group` are still unreached, as before; no
  fixture holds them. The Python CRUD mirror carries the same seed.
- Fidelity (`results/0289/fidelity.*`). The Numbers structural round trip
  adds a table at the start, inserts a row at the start of the original
  table (skipped when that table refuses, typed), then reverts both.
  Statuses are unchanged (38 value-exact, 1 one-way, 1 refused); only the
  label changed.
- Capability and registry goldens are re-blessed for both feature sets, and
  so is the OfficeCLI delta table. The OfficeCLI corpus check passes
  unchanged.

## 6. Performance

The examples are `numbers_unified_bench`, `keynote_bench` and `pages_bench`
(new scenarios are documented in each header). Each was built with
`cargo build --release -p litchi-ops --features <format> --example <bench>`
at `4c611be0e5` and run with `taskset -c 7`. Numbers used 8 ABBA rounds,
and Keynote and Pages used `--rounds 8 --iterations 30`. Every new pair's
outputs are compared byte for byte before timing, and all are identical.

Evidence is in `results/0309/{numbers,keynote,pages}-{1,2}.json[l]`. The
binaries' SHA-256 values: numbers `f2f34064…dabb`, keynote
`dc00af32…4b7c`, pages `03a85011…46ca` (not committed).

The bar is 5% or 20 µs, whichever is larger. Medians from run 1, with run 2
in parentheses:

| Format | Scenario | Leaf | Unified | Ratio | Bar |
|---|---|---:|---:|---:|---|
| Numbers | `row_insert` | 522 µs | 525 µs | 1.005 (1.007) | met |
| Numbers | `row_remove` | 1026 µs | 1036 µs | 1.010 (1.007) | met |
| Numbers | `col_insert` | 488 µs | 490 µs | 1.005 (1.010) | met |
| Numbers | `col_remove` | 1190 µs | 1196 µs | 1.005 (1.001) | met |
| Numbers | `cell_font` | 326 µs | 326 µs | 0.999 (0.997) | met |
| Keynote | `run_style` | 3.76 ms | 3.87 ms | 1.026 (1.024) | met |
| Keynote | `placeholder_hide` | 3.00 ms | 3.09 ms | 1.027 (1.028) | met |
| Keynote | `media_remove` | 7.23 ms | 7.29 ms | 1.009 (1.010) | met |
| Keynote | `table_cell` | 3.07 ms | 6.12 ms | 1.990 (1.987) | **miss** |
| Pages | `paragraph_font` | 3.30 ms | 3.32 ms | 1.007 (1.006) | met |
| Pages | `table_cell` | 1.60 ms | 1.70 ms | 1.065 (1.062) | **miss** (+103 µs) |

The callgrind profiles (`valgrind --tool=callgrind`; this host blocks
`perf`) trace both misses to building the unified tree on the first read,
not to the write:

- **Keynote `table_cell`.** Resolving `…/tr[1]/tc[1]` builds the slide's
  part (`Iwork::kids` → `build_whole`, 222M of 1487M Ir). That build decodes
  the slide table (`select_table_after_inventory`) to make its `tr` and `tc`
  items, which costs about as much as the leaf write. Follow-up: a lazy
  table part for Keynote slide tables (the record 0305 pattern for Numbers
  cells), or reuse of the decoded table between the tree and the write.
- **Pages `table_cell`.** The unified path's only extra work is the
  first-read tree build (`resolve` → `checked_tree` → `build_with`, about
  1.4M Ir per call against a 20.5M Ir leaf commit). The leaf work is
  identical in both legs. Follow-up: a lazy body-table part for Pages.

## 7. Gaps left (typed refusals), with follow-ups

- Numbers topology on tables with formulas or incoming references needs a
  port of the formula-AST coordinate shift. Merges, filters, categories,
  pivots and sort rules (columns) under topology are refused, as is
  removing a hidden, bordered or commented axis. Unified row and column
  `height` and `hidden` are not wired (operation gap).
- Numbers cell writer, pre-existing:
  - several corpus files refuse value writes (`InvalidSource` on cell-value,
    chart, shared-model and table-data-list files; `FormulaCache` on the
    formula and relocation files; `Merge` on table-merges);
  - header-column writes are refused (`HeaderNameIndex`);
  - a write is refused while another table holds formulas;
  - comments (request 19) and cross-table formulas remain open.
- Pages: section add and remove (11); run elements in the tree; a
  whole-section text `set` after styling (`DependentContent`);
  anchored-paragraph moves; Markdown body tables (0306 G3); table creation.
- Keynote:
  - placeholder and notes run styles;
  - table creation and movie and audio creation;
  - duplicating slides whose content lives in other components;
  - removing image data and stale thumbnail records;
  - a readable placeholder `hidden`.
- No native-app acceptance: this host has no macOS.
- `litchi-numbers --all-targets` fails clippy with default features only,
  through dead helpers in `tests/table_relocation.rs`. This predates the
  record, and the gate is `--all-features`.

## 8. Gates (final tree, exit codes)

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-numbers -p litchi-numbers-wire -p litchi-iwa-text-wire -p litchi-keynote -p litchi-pages -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features numbers\|keynote\|pages] --all-targets -- -D warnings` (4 runs) | 0 |
| `cargo clippy -p litchi-ods --lib -- -D warnings` | 0 |
| `cargo test -p <crate> --all-features`: numbers-wire 167, iwa-text-wire 48, keynote 1149, pages 691, numbers 1072 | 0 |
| `cargo test -p litchi-ops --all-features`, chunked: lib 918, doc 210, 137 test targets in 4 chunks (1248) | 0 |
| `cargo test -p litchi --all-features` (230); `cargo test -p litchi-cli --all-features` (33) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p` (changed crates, litchi-ops, litchi) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --profile` numbers-minimal, numbers-full (32 families), pages (26), keynote-minimal, keynote-full (41), ops-iwork | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --test capabilities --test registry` with `--all-features` and with every feature except `encryption`; `officecli_capability_delta.py`; `officecli_corpus.py --check` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` | 0 |
