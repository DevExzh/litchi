# 0312 — Lazy iWork tables, direct cell writes, Pages sections and Keynote placeholder styles

- Status: implemented on `uo2/iwork-tables`, cut from trunk
  `feat/unified-ops-api` at `8f5569d0d8`. Trunk `12724da2d5` (record 0311)
  was merged at `60164238c3`; the merge touched no iWork source.
- Date: 2026-10-16
- Inputs: the two performance misses of record 0309 §6 (Keynote and Pages
  `table_cell`), then the element gaps in the order the coordinator set:
  Pages section add and remove (request 11), Numbers cell comments
  (request 19), Keynote and Pages table creation, styles on Keynote
  placeholder and notes runs, a readable Keynote placeholder `hidden`,
  Numbers row and column edits on tables with formulas; and the
  `litchi-numbers` default-feature clippy failure.
- Owner decisions applied: **D1**, **D2** (2026-10-15): the iWork crates are
  edited directly.
- Commits: `0ba3944a0a` (lazy tables, direct cell writes), `22e5cd03ca`
  (numbers clippy), `b996a70792` (Pages `section::lifecycle`),
  `dd97071e70` (Keynote placeholder styles), `35599e1978` (unified wiring,
  goldens, checklist), `878c09d676` (requests), `60164238c3` (trunk
  merge), `fa5a2415c2` (iwa-archive clippy), `92693c1292` (test), and this
  record.
- `performance_claim`: paired leaf/unified medians in §2 (release builds,
  one pinned core of a shared host, ABBA, two runs). No frozen-harness
  qualification.

## 1. Summary

| Item | Before | After | Status |
|---|---:|---:|---|
| Keynote `table_cell` (unified/leaf) | 1.990 / 1.987 (0309) | 1.018 / 1.014 | **fixed** |
| Pages `table_cell` | 1.065 / 1.062 (0309) | 1.017 / 1.019 | **fixed** |
| Every other 0305/0309 iWork scenario | — | §2 | unchanged; the three pre-existing batch misses stay |
| Pages section add and remove (request 11) | typed gap | `section::lifecycle`; unified `add`/`remove` of `section` | **done** |
| Keynote placeholder run styles | typed gap | title and body placeholders | **done**; notes runs open |
| Keynote placeholder `hidden` | set-only | readable | **done** |
| Numbers cell comments (request 19) | typed gap | diagnosed (§5) | open |
| Keynote and Pages table creation | typed gap | unchanged | open |
| Numbers topology on tables with formulas | typed gap | unchanged | open |
| `litchi-numbers` default-feature clippy | fails | passes | **fixed** |

Reads stay byte-identical except the new placeholder `hidden` property
(§3.3). The query-corpus differential, corpus fidelity and the existing
checklist rows are unchanged; the checklist gains two `ok` cells (Pages
`section` add and remove).

## 2. Lazy table contents and direct cell writes

### 2.1 What the misses were

Callgrind of record 0309's binaries (`--only table_cell`) put both misses
in the first-read tree build:

- **Keynote.** Resolving `/slide[1]/table[1]/tr[1]/tc[1]` built the
  slide's part, which read the slide table (`slide_table_cells`). That read
  is dominated by its package-wide safety census
  (`select_table_after_inventory` → `validate_global_inbound_references`,
  `validate_package_metadata`): about 36 M instructions a call, as much as
  the leaf commit itself. The unified call was resolve 32 M + commit 36 M
  against the leaf's 36 M.
- **Pages.** The first read built the whole tree, body tables included:
  1.39 M instructions a call, most of it `body_table_cells`.

### 2.2 Change (`format/iwork/part.rs`, `keynote.rs`, `pages.rs`, `tree.rs`)

- A Keynote slide table and a Pages body table are now **stand-ins** in the
  tree that holds them. The stand-in carries the table's kind, ordinal,
  path, leaf positions and the cheap properties (frame, drawable or
  catalog name). Its contents (the table item as the reader sees it, with
  `name`, `rows` and `cols`/`columns`, and every row and cell) are built
  whole the first time a read reaches the table's properties or children
  (`part::Contents`, `keynote::build_table`, `pages::build_table`).
  `Iwork::item` serves a stand-in's reads from the built contents;
  `Iwork::stand_in` reads the tree's own item, and kinds, ordinals and the
  write helpers that only need leaf positions use it, so locating a table
  never reads it.
- A Pages table's contents are a part of their own. A Keynote slide's
  tables are discovered only when the slide's part is built, so their
  items live in the slide's part at flagged indices (`part::TABLE`: bit
  31, ten bits of table, 21 bits of item). A table with more than 2^21
  items, or a slide with more than 1,024 tables, keeps the later tables'
  frames without rows.
- **Direct cell writes.** A one-value `set` of `text` or `value` on the
  canonical path `/slide[N]/table[M]/tr[R]/tc[C]` (Keynote) or
  `/body/tbl[N]/tr[R]/tc[C]` (Pages; `table` for `tbl`) locates the
  table through its stand-in and commits the leaf cell write without
  reading the table; the leaf writers check the cell against the table
  (`InvalidSpan`, `InvalidCell`). Any other target or spelling, and any
  failure, takes the general route, which reports it (the record 0305
  DOCX pattern).

Tests:

- `format::iwork::part::tests`: a Keynote slide's other children leave its
  table unread; a kind query reads the stand-in only; one cell builds the
  table once; a direct write builds nothing; a view completes the table.
  The Pages test checks the same for two body tables.
- `tests/iwork_table_cells.rs` compares the direct route with the general
  route (the same cell spelled with the `row`/`cell` aliases, which the
  direct route does not take) on every corpus fixture holding a table: the
  first, last and one-past cells, every value kind, a table past the last.
  Keynote: 36 cells compared, 21 published. Pages: 264 compared, 125
  published. Outputs are equal byte for byte; refusals have equal codes.
- A scratch dump (not committed) of every Pages, Keynote and Numbers
  fixture under `test-data/` (128 packages: all four views as text and
  JSON, a full child walk with kind, path, text, properties and child
  counts, a fresh-handle `get` of every walked path, ten table probes, ten
  queries; 1.25 MB) is byte-identical at the base and after §2.2. After
  §3.3 its only difference is the new `hidden` on the 120 placeholder lines
  (17 of them `true`).

### 2.3 Measurements

Built at `35599e1978`, each bench with only its format feature
(`cargo build --release -p litchi-ops --features <format> --example
<bench>`), run twice with `taskset -c 7`: Keynote and Pages `--rounds 8
--iterations 30`, Numbers `LITCHI_NUMBERS_BENCH_SAMPLES=8`. Evidence:
[results/0312/](results/0312/) (`measure.sh`, `summary.txt`, raw outputs;
binary SHA-256 in `summary.txt`, binaries not committed). Ratio is
unified/leaf, run 1 (run 2); the bar is 5 % or 20 µs, whichever is larger.

| Format | Scenario | Leaf | Unified | Ratio | Bar |
|---|---|---:|---:|---:|---|
| Keynote | `open` | 1625 µs | 1611 µs | 0.992 (0.990) | met |
| Keynote | `open_get` | 1621 µs | 1689 µs | 1.042 (1.034) | met |
| Keynote | `set` | 4967 µs | 5018 µs | 1.011 (1.011) | met |
| Keynote | `notes` | 2786 µs | 2875 µs | 1.033 (1.032) | met |
| Keynote | `frame` | 2972 µs | 3032 µs | 1.021 (1.015) | met |
| Keynote | `batch_1000` | 7956 µs | 8499 µs | 1.068 (1.067) | miss (pre-existing, 0305: 1.071–1.074) |
| Keynote | `duplicate` | 3369 µs | 3466 µs | 1.029 (1.029) | met |
| Keynote | `add_slide` | 16322 µs | 16352 µs | 1.002 (1.002) | met |
| Keynote | `add_textbox` | 3082 µs | 3176 µs | 1.031 (1.032) | met |
| Keynote | `remove_slide` | 6107 µs | 6167 µs | 1.010 (1.011) | met |
| Keynote | `run_style` | 3784 µs | 3902 µs | 1.031 (1.029) | met |
| Keynote | `placeholder_hide` | 3029 µs | 3109 µs | 1.026 (1.032) | met |
| Keynote | **`table_cell`** | 3086 µs | 3142 µs | **1.018 (1.014)** | **met** (was 1.990) |
| Keynote | `media_remove` | 7342 µs | 7436 µs | 1.013 (1.012) | met |
| Pages | `open` | 921 µs | 906 µs | 0.986 (0.995) | met |
| Pages | `open_get` | 923 µs | 937 µs | 1.015 (1.024) | met |
| Pages | `set` | 2017 µs | 2035 µs | 1.010 (1.011) | met |
| Pages | `add` | 2022 µs | 2063 µs | 1.020 (1.019) | met |
| Pages | `remove` | 2013 µs | 2038 µs | 1.012 (1.016) | met |
| Pages | `move` | 2012 µs | 2033 µs | 1.011 (1.012) | met |
| Pages | `paragraph_font` | 3315 µs | 3337 µs | 1.008 (1.014) | met |
| Pages | **`table_cell`** | 1616 µs | 1644 µs | **1.017 (1.019)** | **met** (was 1.065) |
| Pages | `batch_1` / `_10` / `_100` | 2194–2239 µs | 2228–2293 µs | 1.014–1.025 | met |
| Pages | `batch_1000` | 2283 µs | 2474 µs | 1.083 (1.077) | miss (pre-existing, 0305: 1.087–1.098) |
| Pages | `batch_10000` | 2986 µs | 4614 µs | 1.549 (1.527) | miss (pre-existing, 0305: 1.568–1.573) |
| Numbers | `open_get` | 97.5 µs | 102.2 µs | 1.047 (1.039) | met (+4.6 µs) |
| Numbers | `batch_1000` | 960 µs | 1278 µs | 1.332 (1.319) | miss (pre-existing, 0305: 1.315–1.316) |
| Numbers | the 15 other scenarios | — | — | 0.976–1.032 | met |

Every new pair's outputs are compared byte for byte before timing
(`outputs_identical: true` for every Numbers pair; the Keynote and Pages
harnesses assert equality before timing).

**Cost of §3.3 on Keynote reads.** The readable placeholder `hidden` calls
`slide_placeholder_visibility` once per title and body placeholder when a
slide's part is built: callgrind (`--only open_get`) counts 0.48 M
instructions per slide-part build against record 0305's 0.31 M. `open_get`
moved from 1.023–1.034 (0305) to 1.034–1.042, inside the bar. A follow-up
can read both roles in one focused slide selection.

## 3. Element gaps

### 3.1 Pages section add and remove (request 11)

`litchi_pages::section::lifecycle::{Edit, Patch, Commit, Diagnostics,
Error, LimitKind}` with `Package::{edit_section_lifecycle,
apply_section_lifecycle}` and `Edit::{add(after, Option<&str>),
remove(section)}`, one operation per edit (the `sheet::lifecycle` shape).
It ports the retired host editor's `insert_section`/`remove_section` onto
exact-source staging:

- **Add** inserts a U+0004 break at the end of the `after` section and a
  section-table entry (`table_section`, field 17) just after it for an
  independent copy of that section's private graph: the section, its page
  templates, their header and footer storages, the user guide map
  (renumbered to the copy's first page) and its first guide storage
  (emptied), with fresh identifiers in the native editor's order and fresh
  UUIDs for the records that had them. Shared records (styles) stay shared.
  The new section is empty; `name` names it.
- **Remove** drops a non-initial section's table entry and break, and every
  object only its graph owned; the section's text joins the preceding
  section's, as in the native editor.
- The private graph is the closure over the archive headers' reference
  lists (records whose every referrer is in the graph), with the root and
  the registry protected. Refusals: `InitialSection`, `EmptySection` (a
  break must follow content), `UnsupportedContent` (a header or footer
  owning records of its own, data references, unknown record types),
  `UnsupportedDependency` (a surviving object still references the
  graph), `UnsupportedSource`.
- A small structural toolkit (`package/structure.rs`: working set,
  reference census, registry transition, exact artifacts) follows
  `litchi-numbers`' record 0301 machinery. Only the body component and the
  registry are rewritten; every other member keeps its bytes, and the
  inverse restores the source exactly.
- Differential against the retired host editor (`litchi-iwa` examples
  `append_pages_section`, `remove_pages_section`): on `basic.pages` (add)
  and `body-sections-unicode.pages` (remove) the rewritten
  `Index/Document.iwa` is identical to the host's, decompressed; only
  `Index/Metadata.iwa` differs (fresh UUIDs; the host also lowers the
  identifier watermark on removal).
- Tests (`crates/litchi-pages/tests/section_lifecycle.rs`): add with page
  masters (the new section's headers and footers equal its source's),
  reopen, exact inverse, remove again; removal with joined text and exact
  inverse; typed refusals; 31 of 31 corpus documents add a section after
  their last and remove it again. Example:
  `examples/edit_section_lifecycle.rs`.
- **Unified.** `add` of `section` under `/` takes a placement among the
  sections (end by default; `before`/`after` a section, or a position; the
  first section always begins the body) and `name`; `remove` of
  `/section[N]` (sections go last in a mixed removal, from the last). The
  two operation gaps are gone; capability and registry goldens (both
  feature sets), the OfficeCLI delta table and the CRUD checklist are
  re-blessed.

### 3.2 Keynote placeholder run styles

`slide::drawable::Edit::set_text_style` accepts title and body
placeholders: their slide's own storage takes the same character-style
variations as a text box (78 corpus placeholders styled, read back,
inverted exactly; `tests/slide_drawable.rs`). Unified `run` properties
(`bold`, `italic`, `underline`, `size`, `color`, `font.latin`) now write on
placeholder runs too; a batch that writes a placeholder's text before its
run applies them in order. **Open:** speaker-notes runs: the tree lists no
runs under notes, and the notes storage has no run reader or style writer
in the leaf. The operation gap names only notes runs now.

### 3.3 A readable Keynote placeholder `hidden`

`hidden` on `placeholder` is readable: the slide part reads
`Package::slide_placeholder_visibility` for each title and body
placeholder. No leaf change. Test: `placeholders_are_hidden_and_shown`
reads `false`, `true` after the hide, `false` after the show.

## 4. Cleanup

- `litchi-numbers --all-targets` with default features: the three
  helpers of `tests/table_relocation.rs` that only serve tests behind
  `internal-iwork-source` are gated on it (`22e5cd03ca`).
- `litchi-iwa-archive` with default features: `Component::into_parts`'
  one-arm match is allowed there (`fa5a2415c2`); `litchi-pages` and
  `litchi-keynote` default-feature clippy depended on it.

## 5. Not done, with what was found

- **Numbers cell comments (request 19).** `set_table_cell_comment` on every
  corpus workbook (a scratch probe, not committed): 31 files refuse at
  `comments_create.rs` because the creation census is package-wide and
  refuses a workbook with more than one comment list, while native Numbers
  gives each table its own. Selecting the target table's list (tried, then
  reverted) moves the refusal to the next check, which requires the comment
  list, the table model and the tile in one component; native files keep
  tiles in their own components, so creation needs cross-component registry
  references. Six more files fail the strict header-metadata census of the
  other lists (`validate_comment_message_metadata`). Follow-up: a
  per-table census, cross-component creation with external references, and
  a census of the selected list only.
- **Keynote and Pages table creation.** No leaf transaction; not started.
- **Numbers topology on tables with formulas.** The formula-coordinate
  shift is not ported; `Unsupported::{Formulas, ReferencedByOtherOwners}`
  stay.
- **Speaker-notes runs** (§3.2).
- The three batch misses of §2.3 are unchanged (records 0302, 0305).
- No native-app acceptance: this host has no macOS.

## 6. Gates (final tree, exit codes)

Build environment per the rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`,
one cargo command at a time, the memory-capped cargo shim, test executables
pruned between chunks). Run on `92693c1292` (after the trunk merge).

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-pages -p litchi-keynote -p litchi-numbers -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings`; `-p litchi-iwa-archive --all-features` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features keynote\|pages\|numbers] --all-targets -- -D warnings` (4 runs) | 0 |
| `cargo clippy -p litchi-numbers\|litchi-pages\|litchi-keynote\|litchi-iwa-archive --all-targets -- -D warnings` (default features) | 0 |
| `cargo test --all-features -p`: litchi-pages 696, litchi-keynote 1150, litchi-numbers 1072, litchi-iwa-archive 200 | 0 |
| `cargo test -p litchi-ops --all-features`, chunked: `--lib` 920, `--doc` 210, 138 test targets in 4 chunks (1252) | 0 |
| `cargo test -p litchi --all-features` (230); `cargo test -p litchi-cli --all-features` (33, the OfficeCLI corpus included) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-pages -p litchi-keynote -p litchi-numbers -p litchi-iwa-archive -p litchi-ops -p litchi` | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify`; `python3 tools/check_iwork_leaf_public_api.py` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target/glossary` (40 profiles, 2,306 bans, 228 families; `pages` 27) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (279) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |
| `cargo test -p litchi-ops --test capabilities --test registry` with `--all-features` and with every feature but `encryption` (re-blessed in `35599e1978`); `officecli_capability_delta.py` (no change after); `officecli_corpus.py --check` | 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test crud_matrix --test corpus_fidelity` (checklist: two new `ok` cells; fidelity unchanged) | 0 |

## 7. Not claimed

- That the Keynote `batch_1000`, Pages `batch_1000`/`batch_10000` and
  Numbers `batch_1000` misses moved.
- That the direct cell route equals the general route outside the corpus
  (it falls back on every failure, and the corpus comparison is exhaustive
  over the probed cells; a table the reader refuses but the writer
  accepts would write where the general route reports a missing path).
- Native-app acceptance of the new sections and styles.
