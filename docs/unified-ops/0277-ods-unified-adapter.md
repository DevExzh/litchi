# 0277 — Unified ODS adapter (read + write)

- Status: implemented on `uo2/s23-ods`, cut from `feat/unified-ops-api` at
  `395d0ffa9f`. Plan [0248](0248-coordinator-restart-plan.md) §6 Phase F,
  step 23 (ODS stream); builds on the W3 owner split of
  [0268](0268-odf-w3-package-split.md).
- Date: 2026-10-13
- Scope: `crates/litchi-ops/src/format/ods*` (new adapter), small separate
  hunks in the shared dispatch files, `crates/litchi-ods` (leaf gap
  closures), the ops/CLI/Python tests, the capability, registry and query
  goldens, the OfficeCLI delta table and one glossary row. No iWork crate,
  no sibling ODF adapter file and nothing in `litchi-odf-common` changed.
- `performance_claim`: paired leaf/unified medians in §6 (release, one
  pinned CPU, ABBA). No frozen-harness qualification.

## 1. What changed

Before this record ODS opened and read text, counts and metadata only;
`get`/`find`/`query`/`view` were refused with "M1 R7 adapter has not
landed", `capabilities()` was a whole-format refusal, and every write verb,
`save` and `create` returned `unsupported_capability`.

ODS now has a full read adapter and a write adapter at the level of the
other workbook adapters, over the immutable `litchi_ods::Package` and its
`Edit`/`Commit` (record 0268):

| Verb | ODS |
|---|---|
| open | packaged `.ods` (byte, reader and path routes) and flat `.fods` (new; read-only) |
| get / find / query / view | workbook, sheet, cell, range, row, column, named range/expression, annotation (`comment`), merge span (`mergeCell`); `text`, `annotated`, `outline`, `stats` views |
| set | cell `value`/`text`/`clear`/`formula`/`style`; range broadcast and `merge`; sheet `name`; named range `name`/`ref`; annotation `text`/`author`/`date`; `/metadata` |
| add | sheet (any placement), row, named range, annotation |
| remove | cell/range content, sheet, row, named range, annotations (query), merge span |
| move / swap | sheets; rows |
| batch / edit | consecutive direct cell content sets coalesce into one leaf commit |
| save / write_to / to_bytes | the package's exact bytes through `litchi_core::atomic` (sibling temp + rename; `Overwrite::Refuse` honoured) |
| create | one empty `Sheet1`, optional initial `meta.xml` metadata; deterministic |

The CLI and the Python binding pick the adapter up through `litchi-ops`
without code changes; their tests gained one ODS round trip each and the
Python refusing set lost ODS (`READ_ONLY`, `CREATABLE`, flat-ODF fixtures).

## 2. Read adapter (`format/ods/{nav,node,children,index,views,spec}.rs`)

- **Kinds and paths.** The shared workbook vocabulary
  (`format::workbook::spec::Kind`, now compiled for `ods`): `/sheet[N]`,
  `/Sheet1`, `/Sheet1/B3`, `/Sheet1/A1:C5`, `/Sheet1/row[3]`,
  `/Sheet1/col[B]`, `/namedrange[N|name]`, `/Sheet1/comment[N]`,
  `/Sheet1/mergeCell[N]`, and Excel notation. Canonical paths are
  positional (`/sheet[1]/B3`), as in the XLS/XLSB adapters. `slicer` is
  `unsupported_path` (OpenDocument has none); the other unprojected
  elements are typed element gaps.
- **Nodes borrow the owner.** A cell node holds the leaf `Cell` run, a
  sheet node the `Sheet`; nothing is copied per node.
- **Content cells.** ODF stores runs (`number-rows-repeated`,
  `number-columns-repeated`) and producers pad every sheet to 1,048,576 ×
  16,384 with styled empty runs. Queries, row/column listings and views
  enumerate *content cells* only (a value, text, formula, link or merge
  span); padding is skipped without expansion. A per-snapshot, per-sheet
  index (`index.rs`) is built on first use and cached (bounded at 4,194,304
  content cells; `limit_exceeded` beyond). A direct coordinate resolves
  against the typed runs without building an index, so a styled blank still
  answers `get("/Sheet1/Z99")` (as an empty cell). Recorded as the
  `query cell styled-blank-cells` gap.
- **Properties.** Cell: `type` (`Number`/`String`/`Boolean`/`Date`/`Time`/
  `Currency`/`Percentage`/`Unknown`), `formula` (the stored OpenFormula
  token, e.g. `of:=SUM([.A1:.A2])`), `cachedValue`, `empty`, `merge`,
  `mergeAnchor`, `link`; the direct style name is the view's `style`
  member and the query `style` key. Workbook: the ten `meta.xml` keys.
  Named range: `name`, `ref`, `scope`. Annotation: `ref`, `author`, `text`,
  `date`. Merge span: `ref`. Range: `merge`.
- **Not projected (typed property gaps):** sheet `hidden`/`visibility`,
  row `height`/`hidden`, column `width`/`hidden`, cell formatting keys and
  `numberformat` (they live in styles the view does not resolve),
  `computedValue`, `workbook.date1904`, `activeTab`.
- **Flat ODS.** `.fods` opens through `litchi_ods::FlatSpreadsheet` (read
  only). Its definitions are not projected by the leaf (empty) and its
  annotations are a typed `unsupported_capability` (gap F2).
- **Views.** Same layout as the XLS/XLSB adapters, over content rows.

## 3. Write adapter (`format/ods/write.rs`, `write/coalesce.rs`, `metadata.rs`)

Each unified call resolves its targets against the immutable snapshot,
admits properties through static registry tables
(`CELL_WRITE_TABLE`, `MERGE_WRITE_TABLE`, `SHEET_WRITE_TABLE`,
`NAME_WRITE_TABLE`, `COMMENT_WRITE_TABLE`), stages the whole call in one
leaf `Package::edit()` and publishes the leaf commit as a new handle; the
input handle never changes and every refusal leaves it untouched.

- **Cells.** One `worksheets(|w| w.set_cells(..))` per sheet (chunked at the
  leaf's 4,096). A scalar `value` replaces any formula and links; `null`
  clears the value and keeps the style; `clear=true` removes value, text,
  formula, links and range source and keeps style and merge role; `formula`
  stores an inert OpenFormula token (a bare or `=`-prefixed expression gains
  `of:=`) and keeps the cached value; `style` names an existing cell style
  (checked through `Package::effective_cell_style`; unknown names are
  `invalid_value`). Exact no-ops return the source byte for byte.
- **Merge.** `merge=true` spans the range from its first cell and clears the
  covered cells' content (ODF covered cells carry none); `merge=false` and
  removing `mergeCell[N]` unmerge.
- **Sheets.** Add appends one fragment (with one empty row, as producers
  author a new sheet) and moves it for other placements; rename copies the
  table fragment under its new name and drops the original; remove, move and
  swap use the leaf's checked splices. Rename and remove refuse a sheet that
  a formula or definition names; the last sheet is not removable.
- **Rows.** Insert, remove, move and swap through the leaf's logical row
  editor.
- **Named ranges, annotations, `/metadata`.** Through the leaf's
  definitions and annotation families and the new `Edit::set_metadata`
  (title, subject, description, keywords, author; other root-metadata keys
  are a typed refusal naming the property).
- **Leaf refusals** (`InvalidFormat` from a staging step on a document that
  already parsed) are re-coded at the write call site
  (`Error::caused_by_argument`, extended to `ods` and
  `UnsupportedCapability`) while keeping the typed leaf source:
  `invalid_value` for authored values, `unsupported_capability` for document
  shapes the leaf will not rewrite.
- **Batches.** `transaction::Edit` gained an ODS cell run
  (`write::coalesce::Run`): consecutive direct single-cell `value`/`text`/
  `clear`/`formula` sets to distinct coordinates share one leaf commit; a
  repeated coordinate, a style or merge set, a range, a query or any other
  verb closes the run. Per-item error attribution and atomicity are those of
  the other coalescers.

## 4. Leaf gap closures (`crates/litchi-ods`, separate hunks)

LibreOffice-produced spreadsheets could not be edited at all through the
leaf's worksheet route before these; each closure keeps every untouched
source byte and is pinned by `tests/producer_part_edits.rs`.

| # | Gap | Closure |
|---|---|---|
| L1 | Row splices refused `calcext:value-type`, which LibreOffice writes on every typed cell | admitted when it restates the cell's `office:value-type`; the row writer re-emits it (binding the namespace on the row) for splices whose source rows carried it (`worksheet/{package,codec}.rs`) |
| L2 | A row splice whose source `content.xml` has a newline after the XML declaration fell back to a whole-table rewrite (refused on producer tables) | only a rebuilt `content.xml` must be compact; spliced fragments stay individually audited (`worksheet/snapshot.rs`) |
| L3 | A cell below the last physical row (or in a row-less table) had no splice anchor | pure insertions splice after the preceding row, before the next, or at the end of a table holding only column declarations (`worksheet/package.rs`) |
| L4 | `Edit::commit` audited an unchanged producer manifest for compactness | only an edited manifest is audited (`document.rs`) |
| L5 | Producer prolog whitespace failed whole-part audits (definitions, annotations, metadata) | prolog whitespace is not authored content: excused by the audits and dropped from whole-part republication (`package/document.rs`, `definitions.rs`, `document.rs`) |
| L6 | A metadata edit republished `styles.xml`/`settings.xml` as authored XML | `replace_metadata_xml` copies every other member exactly (`rebuild_package` with a `meta.xml` addition) |
| L7 | No metadata write on an opened package | `document::Edit::set_metadata(&Metadata)` stages the fields the leaf patcher owns (op `metadata.edit`) |
| L8 | Flat documents with XML `Misc` after the root (every real `.fods`) were refused | whitespace, comments and PIs around the root are admitted; text and CDATA outside it are still refused (test updated) |
| L9 | `Package::new()` built zero sheets although documented as one | it builds one empty `Sheet1` |
| L10 | `SourceBackedSpreadsheet::materialize` consumed the source owner | takes `&self`, so a retained owner materializes its editable package on the first edit |
| L11 | Flat spreadsheets had no text projection | `FlatSpreadsheet::text()` (shared projection; `facade::source` is now `pub(crate)`) |

0268 follow-ups G2–G4 were not needed: the adapter uses only `Package`,
`Edit`, `Commit` and the read projections.

## 5. Capabilities, registry and goldens

- `capabilities()` and `schema()` report ODS per element from the same
  spec and registry tables: 26 elements (9 readable, 17 typed gaps), write
  flags from the write tables (`SETTABLE`, `ADDABLE`, `REMOVABLE`) and 12
  contextual write gaps (`WRITE_GAPS`); import and evaluation are typed
  `unsupported_capability` (OpenFormula evaluation is not integrated).
- New goldens `tests/golden/{capabilities,registry}/ods.json`, identical
  with and without the `encryption` feature; no other golden changed.
- `docs/unified-ops/officecli-capability-delta.md` regenerated
  (`python3 tools/officecli_capability_delta.py --officecli
  /home/zhuhe/code/litchi/3rdparty/OfficeCLI`): the tool already compares
  ODS against the XLSX vocabulary as a family parity target (no OfficeCLI
  ODS schema exists); `--check` against both trees and
  `python3 -m unittest tools.test_officecli_capability_delta` pass.
- Query corpus: `tests/support/query_corpus.rs` gained `ods` (13 files
  opened) and `fods` (13 opened, 2 refused by the leaf flat parser) over the
  XLSX selector set, including the early-stop
  digests. The 224 new digests were blessed from the read adapter before any
  writer existed and reproduced byte-identically by two further runs; every
  existing digest is unchanged.

## 6. Measurements

`crates/litchi-ops/examples/ods_unified_bench.rs`
(`litchi.ods_unified_bench.v1`): one 1,000 × 10 sheet (numbers and text,
43,574 bytes), both legs on the same bytes; write outputs compared byte for
byte first (identical in every scenario). Release build, `taskset -c 7`
(AMD EPYC 9R45, shared host), 5 ABBA rounds (10 samples per leg) per run.
Ratio = unified median / leaf median. Evidence: `results/0277/`.

| Scenario | leaf median (µs) | unified median (µs) | ratio (per-run range) | write outputs |
|---|---:|---:|---:|---|
| `open_get` | 12,402.8 | 12,417.9 | 1.001 (1.000–1.002) | — |
| `query` (all 10,000 cells) | 50.9 | 555.3 | 10.9 (9.7–22.0) | — |
| `set` (one cell) | 62,806.1 | 63,280.9 | 1.008 (0.986–1.012) | identical |
| `batch_1` | 62,463.8 | 62,219.8 | 0.996 (0.973–1.004) | identical |
| `batch_10` | 61,921.1 | 61,865.1 | 0.999 (0.996–1.010) | identical |
| `batch_100` | 62,098.9 | 62,363.7 | 1.004 (1.001–1.014) | identical |
| `batch_1000` | 61,048.7 | 62,352.7 | 1.021 (1.016–1.030) | identical |
| `batch_10000` | 39,754.4 | 56,559.6 | 1.423 (1.395–1.432) | identical |

Medians of the four runs' medians. The binary is built from this record's
commit with `cargo build --release -p litchi-ops --features ods --example
ods_unified_bench` (SHA-256
`b50a30d44efde8b0cd365fce338165722bedda210ffafd14ba5f184a278ec092`) and run
as `LITCHI_ODS_BENCH_SAMPLES=5 taskset -c 7
target/release/examples/ods_unified_bench`; the four JSON reports
(`results/0277/ods-unified-bench-run-{1..4}.json`, about 5 KiB each) hold
every sample. The leaf edit itself dominates every write (a full worksheet
parse, a row splice and a typed readback reparse: about 62 ms here).

- **Open + first get, single set, 1–1,000-op batches: at parity.** The
  single set is within 1% of the 62 ms leaf edit (budget 10%); every batch
  up to 1,000 operations is one leaf commit and stays within 3% of the
  hand-written leaf `set_cells` (budget 2×).
- **10,000-op batch: ~1.4× the leaf** (within the 2× budget). The unified
  run resolves each item's path and plans it against the typed runs
  (`Sheet::cell` scans the physical rows), so planning is O(items × rows)
  on top of the leaf's own O(items × rows) no-op filter. Follow-up P1:
  plan coordinates against the cached sheet index.
- **Query: ~10× a raw loop over the typed runs.** The leaf leg is a bare
  loop summing text lengths (no nodes, no paths); the unified leg builds
  the content index on a fresh handle and yields node views. No budget
  applies to this comparison; recorded for context.

## 7. Gaps and follow-ups

- **F1 flat writes.** `.fods` is read-only; the leaf flat transaction
  admits scalar cell writes only.
- **F2 flat annotations and definitions** are not projected.
- **M1 metadata family** (owner decision 2026-10-14: the current
  `Edit::set_metadata` shape stays; the family is a recorded improvement).
  ADR 0034's `Package::edit_metadata` +
  `metadata::{Edit, …}` family is not built; `document::Edit::set_metadata`
  covers the five owned fields (glossary row annotated). `lastModifiedBy`,
  `created`, `modified`, `language`, revision, template and custom
  properties are typed refusals.
- **S1 styles.** Formatting keys (font, fill, border, alignment, number
  format) are neither read nor written; cells take a named style.
- **S2 layout.** Row height/visibility and column width/visibility.
- **R1 producer settings.** LibreOffice files carry `settings.xml`; the
  leaf's logical row editor refuses row insert/remove/move/swap on them,
  and its sheet copy/move refuse `manifest.rdf`/settings owners, so rename,
  move and swap of sheets succeed on created or settings-free workbooks
  only (typed `unsupported_capability` otherwise).
- **C1 columns.** Column insert/remove (no logical column editor).
- **E1 evaluation and import.** OpenFormula evaluation and CSV import.
- **E2 elements.** Tables (database ranges), charts, pivots (data pilot),
  sparklines, validations, conditional formats, auto-filters, hyperlinks as
  nodes, drawings, page breaks and rich-text runs remain element gaps.
- **P1** as in §6.
- **Duplication for later consolidation.** The ODS views, `TrimEnd` and the
  sheet-name validation mirror the XLS adapter's; `format/odf_ods.rs` was
  not needed (no helper shared with the sibling ODF adapters yet).

## 8. Gates

| Command | Result |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ods -p litchi-ops -p litchi -p litchi-py -p litchi-cli --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-ods --all-features --no-fail-fast` | 0 (1,778 passed) |
| `cargo test -p litchi-ops --all-features --no-fail-fast` | 0 (1,996 passed; query, capability and registry goldens as in §5) |
| `cargo test -p litchi --all-features --no-fail-fast` | 0 (260 passed) |
| `cargo test -p litchi-cli --all-features --no-fail-fast` | 0 (27 passed) |
| `cargo test -p litchi-ops --no-default-features --features "doc docx ppt pptx xls xlsx xlsb rtf odt ods odp markdown formula yaml eval" --test capabilities --test registry` | 0 (goldens without `encryption`) |
| `cargo check -p litchi-ops --no-default-features --features ods --tests --examples` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ods -p litchi-ops` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 (26 profiles, 1,201 bans, 38 families) |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 (254 tests) |
| `python3 tools/officecli_capability_delta.py --check` (vendored and `--officecli` trees) | 0 |
| `python3 -m unittest tools.test_officecli_capability_delta` | 0 (24 tests) |

The `litchi-ods` test target `ods_formula_conditional_native` includes a
5.5 MiB evidence file purged from the repository on 2026-10-05
(`docs/report/spec-gap-validation-evidence/ods-formula-conditional-aggregates/native/cached-results.json`);
it was copied untracked (it is listed in `.git/info/exclude`) from the main
checkout to build the suite and is not committed.

## 9. Not claimed

No evaluation, no styled-cell formatting, no flat writes, no producer-file
row shifts; no performance claim beyond §6.
