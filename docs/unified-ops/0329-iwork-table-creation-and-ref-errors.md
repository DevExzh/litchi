# 0329 — Keynote and Pages table creation, Numbers `#REF!`, shared-formula splits, structural clears, comment edge cases, Keynote notes run styles

- Status: implemented on `uo2/iwork-tables2`, cut from trunk
  `feat/unified-ops-api` at `fc274041da`; trunk merged at `fc2c1a29e1`
  (`d761892de3`, record 0330), `2fce2b639d` (`3bec513e6a`, records 0327
  and 0328) and `99d6560eef` (`0056b9cccd`, record 0332).
- Date: 2026-10-17
- Inputs: records 0298, 0299, 0301, 0309, 0312 and 0318 and
  `IWORK_API_REQUESTS.md` requests 6, 10 and 19. The work list: Keynote and
  Pages table creation; Numbers `#REF!` for deleted references; splitting
  shared formulas whose cells diverge; the clears a row or column removal
  needs; the comment edge cases record 0318 left (empty cells in unstored
  rows, several authors, a list in a member of its own); Keynote
  speaker-notes run styles.
- Owner decisions applied: **D1**, **D2** (2026-10-15): the iWork crates
  are edited directly.
- Owned: `litchi-numbers-wire`, `litchi-keynote`, `litchi-pages`,
  `litchi-numbers`, the `litchi-iwa` differential test,
  `litchi-ops/src/format/iwork/**`, the iWork rows of the goldens, the CRUD
  checklist and the delta table, `IWORK_API_REQUESTS.md`, the iWork bench
  routes.
- Commits: `e0ff6fe996` (shared table payload builders), `46fbdfa1cd`
  (Keynote slide-table creation), `533b3b227d` (Pages body-table
  creation), `25438a113e` (unified `add` of tables), `3457b067b2` (BNC
  error caches and re-keying), `aed6df77d5` (`#REF!`, splits, structural
  clears, differential test), `f17666d1cb` (unified removals under
  formulas), `ed77f468f7` (comment edge cases), `5209b3cbd0` (Keynote notes
  run styles), `41bdfc1cca` (unified notes runs), `f8cada4faf` (bench
  routes), the trunk merge `fc2c1a29e1`, `276e8d431b` (the `iwork_unified`
  expectation, §3.1), and this record.
- `performance_claim`: paired leaf/unified medians in §7. Evidence:
  [results/0329/](results/0329/) (`measure.sh`, `summary.txt`,
  `summarize.py`, raw outputs, `binaries.sha256`; binaries not committed).

## 1. Summary

| Item | Status |
|---|---|
| Keynote slide-table creation | **done** (§2): leaf `Edit::add_table`, unified `add` of `table`/`tbl`; a deck without a slide table refuses, typed |
| Pages body-table creation | **done** (§3): leaf `Package::insert_body_table`, unified `add` of `tbl`; a document without a body table refuses, typed |
| Numbers `#REF!` on deleted references | **done** (§4.1): single references and wholly removed ranges become `REFERENCE_ERROR_NODE`, ranges that lose an edge row or column shrink |
| Splitting shared formulas | **done** (§4.2) for unsegmented formula lists |
| Structural clears on removal | **done** (§4.3): comments through the comment transaction, everything else dropped by the structural pass when the cell writer cannot clear it |
| Comments on empty cells of unstored rows | **done** (§5) |
| Comments among several authors | **done** (§5) |
| A new comment list in a member of its own | **not done** (§5.3): needs a registry component addition |
| Keynote speaker-notes run styles | **done** (§6) |

Existing read paths and existing write results are byte-identical: the
query-corpus differential digests change only for the new notes runs
(one key renumbered, `run` 106 → 108 nodes), corpus fidelity passes
unchanged, and every existing CRUD checklist cell is unchanged. The
capability and registry goldens (both feature sets), the CRUD checklist
and the OfficeCLI delta table change for the new cells only.

## 2. Keynote slide-table creation (`litchi-keynote`)

- `slide::drawable::NewTable::{new(rows, columns, Geometry), with_name,
  with_cell}` and `Edit::add_table(slide, NewTable)`.
- **Template.** The deck's first slide table in slide and z-order (read
  once per edit). The new table takes the template's table drawable and
  model and one emptied copy of each storage object the model owns: tiles
  hold no rows, header buckets no headers, data lists no entries, the
  merge map no ranges, and the row/column UID map and stroke sidecar are
  rebuilt for the requested size. Styles stay shared. What the template
  owns beyond its storage (sort, filter, category, pivot and hidden-state
  owners, summary model, title, caption, comment) is left out. The table
  gets its own formula owner in the calculation engine; the registry gets
  the new object UUIDs and cross-member references. Rows and columns
  divide the frame evenly; the name defaults to `Table N` after the deck's
  tables. Fresh identities derive from seeds, not a clock.
- `litchi_numbers_wire::table_create` is the payload half shared with
  Pages: `Shape`, `admits_size`, `TableUuid::derive`, `model_payload`,
  `info_payload`, `empty_storage`, `owner_payload`, `engine_payload` and
  the reference builders.
- Seeded cells go through the shared cell writer (records 0309, 0312), the
  table is read back through `slide_table_cells` before publishing, and the
  patch inverts to the source bytes.
- Refusals: `Unsupported::TableTemplate` (no slide table to copy),
  `Error::InvalidTableSize` (empty, past the identity ceiling, a seeded
  cell outside the table or a non-finite number), `InvalidGeometry`,
  `LimitExceeded`, and the drawable family's registry and foreign-member
  refusals.
- Corpus (`tests/slide_table_creation.rs`): 6 decks create a table that
  reads back, reopens and inverts exactly; 104 refuse typed (53 of the 56
  corpus decks hold no slide table).

## 3. Pages body-table creation (`litchi-pages`)

- `Package::insert_body_table(anchor: Position, table::creation::NewTable)`
  → `table::creation::Commit {package, table, patch}`;
  `apply_body_table_creation(&Patch)` replays a patch or its inverse.
- The new table copies the document's first body table the way Keynote's
  does (§2), plus its drawable attachment; it is anchored with a `U+FFFC`
  at a UTF-16 body position and a body attachment-table entry, later text
  attributes shifting as a native text insertion shifts them. No clock or
  randomness is read: the same insertion publishes the same bytes. Seeded cells are written through
  `table::values`.
- The structural toolkit's registry transition (`package/structure.rs`)
  gained external-reference additions and component lookup by member.
- Refusals: `Error::NoTemplate` (no body table), `Error::Unsupported` (the
  body-table catalog refuses the document), `InvalidSize`, `InvalidCell`,
  `PositionOutOfBounds`, `LimitExceeded`.
- Corpus (`tests/body_table_creation.rs`, anchors at the start and at the
  end of the first section): 34 insertions read back, reopen and invert
  exactly; 16 refuse typed (14 documents hold no body table).

### 3.1 Unified `add`

- Keynote: `add` of `table` or `tbl` under `/slide[N]` with `rows`, `cols`
  (or `columns`), `name`, `data` (`"a,b;c,d"`, rows by `;`) and the frame
  keys (`x`, `y`, `width`, `height`; default 100, 100, 600 × 200 pt).
- Pages: `add` of `tbl` under `/body` with `rows`, `cols`, `name`, `data`;
  an empty paragraph is added at the placement and the table anchored in
  it, so the table is its own body child.
- Error codes: `InvalidTableSize`, and Pages' `InvalidSize`, `InvalidCell`
  and `PositionOutOfBounds` → `out_of_range`; `LimitExceeded` →
  `limit_exceeded`; `TableTemplate`/`NoTemplate` →
  `unsupported_capability`, with the gap contexts `a deck without a slide
  table` and `a document without a body table`.
- Goldens: the Keynote and Pages capability and registry rows mark
  `table`/`tbl` addable with the new keys; the CRUD checklist gains
  `iwork/keynote/slide-table-read-native.key + add table` and
  `iwork/pages/body-table-read-native.pages + add table`.
- One existing expectation changed on purpose: `iwork_unified.rs` asserted
  that `add` of `tbl` on `basic.pages` is `unsupported_capability`; it now
  adds `rows`/`cols` under `/body` and gets the same code from
  `NoTemplate` (the document has no body table).

## 4. Numbers row and column removal under formulas (`litchi-numbers`)

### 4.1 `#REF!`

Removing a row or column that a surviving formula reads no longer
refuses (`DeletedReference`):

- a single reference to a removed cell, or a range wholly on removed
  rows or columns, becomes `REFERENCE_ERROR_NODE` (AST node type 30,
  prints `#REF!`); the formula's cached result becomes an error
  (`BncCell::set_formula_cached_error`, cell type 8, its string key
  released);
- a range that loses an edge row or column shrinks; mixed absolute and
  relative endpoints are rewritten per endpoint;
- precedent edges and range dependencies into the removed cells are
  dropped; a dependency tile left empty is removed, as the host editor
  does.

Still refused (`DeletedReference`): several ranges of one tract with one
of them removed, and merge-range endpoints.

### 4.2 Shared-formula splits

A formula-list entry several cells share, whose cells would diverge after
the rewrite, is split into one entry per distinct formula
(`BncCell::set_formula_identifier` re-keys the cells; a freed identifier
is reclaimed into `nextListID`). A split in a segmented formula list
still refuses (`SharedFormulas`).

### 4.3 Structural clears

- Comments on a removed row or column are removed through the comment
  transaction first.
- Where the cell writer cannot clear the removed cells (`HeaderNameIndex`,
  `FormulaCache`, `InvalidSource`, the created workbook's reference
  budget), the structural pass drops them with their strings, formulas
  and dependency records. A range over such cells that would keep a stale
  cache refuses (`Formulas`).

### 4.4 Differential test

`crates/litchi-iwa/tests/numbers_topology_formula_differential.rs` now
checks removals against the retired host editor and a text oracle that
writes `#REF!` and shrinks ranges: every removal the host editor
publishes over the differential corpus (18 cases), the leaf publishes
with equal calculation-engine and formula-list content; 10 of them write
`#REF!`. The unit test `deleted_references_cache_an_error` and
`tests/topology.rs::shared_formulas_split` cover the leaf alone.

Unified: `remove` of a row or column a formula reads now publishes
(`numbers_unified.rs` expects `=(SUM($B$2:B2)+#REF!)`); the Numbers row
gap text names the remaining refusals.

## 5. Numbers comment edge cases (request 19)

1. **Empty cells in unstored rows.** `CommentSet` materialises a row its
   tile does not store; the corpus sweep (`tests/comment_corpus.rs`) adds
   a pass over an empty cell of every table (31 tables).
2. **Several authors.** A new comment is the generated author's
   (`litchi-iwa`): reused when the package holds it, else appended to the
   author storage (`append_author_storage`). No present author is
   impersonated. `tests/comment_authors.rs` builds the cases from
   `comment-metadata-native-resaved.numbers` by surgery (a second author
   and its registry edge).
3. **Not done: a comment list in a member of its own.** The new list still
   goes beside the table model. A member of its own needs a component
   registration, and `litchi_iwa_protos::package_metadata_codec` edits
   existing components only. Follow-up: a `ComponentAddition` in its
   batches.

## 6. Keynote speaker-notes run styles (request 10)

- `slide::drawable::Edit::set_notes_text_style(slide, TextSpan, Style)` and
  `Package::slide_notes_text_runs(slide)`; the notes storage is styled with
  the same character-style writer as text boxes and placeholders (record
  0312), the text kept, the inverse exact.
- Refusals: `Unsupported::TextStyling` (a style that sets nothing),
  `InvalidSpan`, and the drawable family's registry refusals.
- `tests/slide_notes_style.rs`: every corpus slide with notes (2 slides,
  none refused); written notes take styles on a sub-span on every slide
  of `multi-slide.key`.
- Unified: `/slide[N]/notes/p[M]/run[K]` reads and `set`s character
  properties; the notes-run operation gap is gone (Keynote operation gaps
  9 → 8).

## 7. Measurements

Release builds of `-p litchi-ops --features <format> --example
<bench>` (`keynote_bench`, `pages_bench`, `numbers_unified_bench`) at
`fc274041da` (base) and `f8cada4faf` (after), `taskset -c 7` on a shared
32-core host (load 5–9), in-process ABBA, processes B A A B
(`results/0329/measure.sh`). Ratio is unified/leaf; every pair's outputs
are byte-identical before timing. Bar: 5 % or 20 µs, whichever is larger.

| Scenario (new) | After 2 / 3 ratio | Leaf / unified µs (after 3) | Bar |
|---|---:|---:|---|
| Keynote `table_add` | 1.010 / 1.012 | 6929.7 / 7010.7 | **met** |
| Keynote `notes_run_style` | 1.071 / 1.033 | 4002.9 / 4134.8 | **met** (+132–289 µs; run 2 had one noisy round, the bar's 5 % arm is 200 µs) |
| Pages `table_add` | 0.995 / 1.005 | 6245.6 / 6274.8 | **met** |
| Numbers `row_remove_ref` | 1.023 / 1.018 | 548.3 / 558.2 | **met** (+10 µs) |
| Numbers `comment_empty` | 1.004 / 1.004 | 4281.4 / 4297.2 | **met** |

- The first Keynote `table_add` measured 1.46: the unified leg built the
  table contents to name the new node's kind. It now counts tables by kind
  ordinal (1.01).
- **Existing scenarios, after against base.** Every ratio is within
  0.02 of base (`summary.txt`) except Keynote `run_style` run 2 (1.043 vs
  1.030; run 3 1.033). The Keynote after processes are 2–5 % slower on
  both legs of every scenario, `open` included, which this record does
  not touch; callgrind instruction counts per call are unchanged
  (`run_style` +0.06 %, `open` +0.01 %), so this is code layout and host
  noise, not added work. Pages and Numbers wall-clock times match base.
- Pre-existing misses unchanged: Numbers `batch_1000*`, Pages
  `batch_10000*` (records 0314, 0326).

## 8. Gaps and follow-ups

1. **Built-in table template**: a deck or document without a table to copy
   refuses (53 of 56 corpus decks, 14 corpus documents). Follow-up: a
   template built from the theme's table style preset and a storage
   skeleton.
2. **Comment list in a member of its own** (§5.3): `ComponentAddition` in
   `package_metadata_codec`.
3. **`#REF!` remainders**: several ranges of one tract with one removed,
   merge-range endpoints, splits in segmented formula lists, UID-reference
   formula nodes (pre-existing).
4. **Drop-mode ranges**: a range over cells the cell writer cannot clear
   refuses rather than leave a stale cache; recomputing the cache would
   need the evaluator.
5. **Native-app acceptance**: this host has no macOS; that a created
   table, a `#REF!` formula, a split formula list or a styled notes run
   opens in Keynote, Pages or Numbers is not claimed.

## 9. Gates (merged tree, exit codes)

Build environment per the rules (`CARGO_BUILD_JOBS=6`,
`RUST_TEST_THREADS=8`, one cargo command at a time, the memory-capped cargo
shim, test executables pruned between chunks).

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` (after the merge, and after the last edit) | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-keynote -p litchi-pages -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features keynote\|pages\|numbers] --all-targets -- -D warnings` (4 runs) | 0 |
| `cargo clippy -p litchi-numbers-wire -p litchi-numbers -p litchi-keynote -p litchi-pages --all-targets -- -D warnings` (default features) | 0 |
| `cargo test --all-features -p` `litchi-numbers-wire` (171), `litchi-keynote` (1,159), `litchi-pages` (702), `litchi-numbers` (1,080), `litchi-iwa` (1,783) | 0 |
| `cargo test -p litchi-ops --all-features`, chunked: `--lib` 929, `--doc` 210, 147 test targets in 10 groups (1,338) | 0 (after the `iwork_unified` expectation in §3.1) |
| `cargo test -p litchi-ops --no-default-features --features` every feature but `encryption` `--test capabilities --test registry` | 0 |
| `cargo test -p litchi --all-features` (230) | 0 |
| `cargo test -p litchi-cli --all-features --test officecli_corpus` (5); `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --officecli …/3rdparty/OfficeCLI --check` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-numbers-wire -p litchi-numbers -p litchi-keynote -p litchi-pages -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/check_iwork_leaf_public_api.py`; `python3 tools/test_no_3rdparty_references.py`; `python3 tools/unified_bench_matrix.py --check` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,313 bans, 252 families) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (283) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

After the last two merges (records 0327, 0328, 0332): `cargo fmt --all
--check`, the all-features clippy above, the capability and registry
goldens in both feature sets, `crud_matrix`, `corpus_fidelity`,
`query_corpus_differential`, `acceptance_fixtures` and the iWork unified
targets (12 targets, 88 tests), the OfficeCLI corpus re-recorded
(`LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
officecli_corpus`, unchanged), and the regenerated delta table and corpus
report with `tools/officecli_dispositions.json`:
`python3 tools/officecli_capability_delta.py --officecli … --check` and
`python3 tools/officecli_corpus.py --check` — all exit 0, with no
dispositions change (iWork formats have no OfficeCLI vocabulary).

Goldens: capabilities and registry (both feature sets), the CRUD
checklist, the query-corpus digests and the delta table were re-blessed in
the feature commits and check clean on the merged tree. The OfficeCLI
script corpus runs OOXML scripts only; its outcomes do not change.
Target directory 17 GB after pruning (26 GB before: test examples and the
glossary rustdoc tree).
