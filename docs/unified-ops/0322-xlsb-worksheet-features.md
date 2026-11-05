# 0322 — XLSB worksheet features (the nodes of record 0320) and XLSX follow-ups

- Status: implemented on `uo2/xlsb-features`, cut from `feat/unified-ops-api`
  at `7ede6ecbe6`, trunk merged at `db96990ff6`.
- Date: 2026-10-17
- Input: record 0320 (the nine XLSX worksheet feature nodes) and its gaps 3
  (threaded comments), 7 (prefixed-root validation writer) and the
  auto-filter `criteriaN.op` row.
- Owned: `litchi-xlsb` (one staged part engine, seven edit families, the
  feature and chart reads); `litchi-xlsx` (small hunks: the prefix fix,
  `criteriaN.op` support, threaded-comment reads); the XLSB adapter of
  `litchi-ops` (`format/xlsb/{features,feature_write,chart}.rs` and small
  hunks in `node`, `nav`, `children`, `spec`, `write`); the XLSX
  `format/xlsx/{features,feature_write}.rs` criteria hunks; the XLSB rows of
  the registry, capability, checklist and fidelity goldens; and one Python
  test.
- `performance_claim`: none beyond §5, which reports paired measurements of
  the new paths.

## 1. Result

| OfficeCLI element | XLSB before | XLSB after |
|---|---|---|
| `comment` (notes) | gap | get, query, add (sheet + `ref`, or a cell parent), set (`text`, `author`, `ref`), remove; VML note shapes kept in step |
| `hyperlink` | gap | get, query, add, set, remove; external targets as worksheet relationships |
| `picture` | gap | get, query, add (`src`, geometry in cells or lengths, `anchor`, `anchorMode`), set, remove |
| `shape` | gap | get, query, add (preset geometry, fill, line, text runs, alignment, margins, rotation, flips), set, remove |
| `conditionalformatting` and its 12 rule kinds | gap | get, query, add, set, remove; formatting through a new BIFF12 differential format (`BrtDXF`) |
| `validation` | gap | get, query, add, set, remove (classic `BrtDVal`; overlap refused, as OfficeCLI does) |
| `autofilter` | gap | get, query, add, set (`range`), remove |
| `rowbreak`, `colbreak`, `pagebreak` | gap | get, query, add, set, remove |
| `chart`, `series`, `axis` | gap | get and query (read only; `axis[@role=…]` resolves, as on XLSX) |

XLSB now matches XLSX, using the same property names and registry entries.

Delta table (`officecli-capability-delta.md`, XLSB row):
- elements supported **8 → 33**, gap **30 → 5** (the same five as XLSX);
- property reads supported **78 → 295**;
- property writes supported **8 → 165**.

The OfficeCLI corpus has no XLSB calls. A re-record left `outcomes.tsv` and
`officecli-corpus-deltas.md` unchanged.

XLSX follow-ups from 0320:

- Gap 7 is fixed. On a worksheet whose root has a prefix (`<x:worksheet>`),
  the validation writer used to emit `<dataValidations xmlns=SML>`, and the
  hyperlink writer then refused the sheet. The writer now emits the
  collection in the root prefix. The parser binds that prefix for core
  rules. Regression test: `a_prefixed_root_keeps_the_worksheet_writable`.
- Auto-filter `criteriaN.op`: `criteria0` is now an add-time property. The
  `criteriaN.OP` spellings (`equals`, `gt`, `top`, `dynamic`, …) build
  `filterColumn` entries the way OfficeCLI does, so `CRITERIA_GAP` was
  removed.
- Threaded comments (gap 3, read only): `{Worksheet, SourceWorksheet}::
  threaded_comments` and `{Workbook, SourceBackedWorkbook}::people`.

## 2. Leaf: `litchi-xlsb`

The crate-private engine `part_edit` is the BIFF12 counterpart of the XLSX
engine:

- `Staged` clones the OPC graph and records the parts it touches.
- `commit` captures the exact before/after `PartPatch`, then publishes by
  reparsing.
- An unchanged staging returns the source snapshot.

Worksheet streams are edited at the record level. `SheetLayout` splits the
records after the cell table into top-level units ranked in MS-XLSB
2.1.7.62 order. A family finds its insertion point by rank, and every record
it does not edit keeps its bytes. `worksheet_index::maintain` runs whenever
a worksheet stream changes.

Seven ADR 0034 families are registered in the glossary gate (profiles
`xlsb-minimal` and `xlsb-full`). Each one starts from a `Package` or,
without a copy, from a `Workbook`:

| family | edits |
|---|---|
| `page_breaks::edit` | `page_breaks`, `set_page_breaks` (`BrtBeginRwBrk`/`BrtBeginColBrk`) |
| `auto_filter::edit` | `auto_filter`, `set_auto_filter`; a moved range keeps the filter columns still inside it |
| `hyperlinks::edit` | `hyperlinks`, `add`, `set`, `remove`; external targets get or release relationships |
| `data_validation::edit` | `validations`, `add`, `set`, `remove`; classic `BrtDVal` inside `BrtBeginDVals` |
| `conditional_formatting::edit` | `formats`, `add`, `set`, `remove`, `differential_formats`, `add_differential`; `DifferentialStyle::{parse, encode}`; priorities must be unique |
| `comments::edit` | `comments`, `add`, `set`, `remove`; authors, `COMMENT` units and VML note shapes |
| `drawing::edit` | `elements`, `add_picture`, `add_shape`, `set`, `remove`; the XLSX drawing vocabulary with `Rgb = u32`; `ChartPart` |

Reads work on both owners and do not decode cells:

- `Workbook::{sheet_features, sheet_comments, drawing_elements,
  drawing_charts, differential_formats}`;
- `SourceBackedWorksheet::{features, comments, drawing_elements,
  drawing_charts}`.

`FeatureSet` selects which kinds a read returns. Only validations and formats
render formulas, so a read that skips them never starts the formula pass.

CF formulas compile with `Compiler::compile_shared` (`PtgRefN`, relative to
the first range's top-left cell). The renderer parenthesizes every binary
operation, so `B2>5` reads back as `(B2>5)`.

## 3. Unified layer: `litchi-ops`

- `format/xlsb/features.rs`:
  - `Item::Feature(Arc<Feature>)` and its readers;
  - property tables copied from the XLSX feature tables, with the XLSB gaps
    marked;
  - `FEATURE_ELEMENTS`, which feeds the capabilities.
- `format/xlsb/feature_write.rs`:
  - add, set and remove for each family;
  - `Patch::Feature(Box<Images>)`, which serializes package images lazily.
- `format/xlsb/chart.rs`: read-only chart, series and axis nodes. The axis
  spec declares the `role` identity, as XLSX does.
- Shared-file hunks:
  - `capabilities.rs`: XLSB settable/add/remove now include
    `FEATURE_ELEMENTS`;
  - `format/mod.rs` and `format/chart.rs`: the chart cfg also covers xlsb;
  - `Cargo.toml`: xlsb gains `dep:base64` and `dep:litchi-drawingml`, plus
    the `xlsb_features_bench` example.

## 4. Tests, checklist, goldens

- `tests/xlsb_features.rs` (10 tests):
  - each kind is added, read back on both owners after a save, set, queried
    and removed;
  - add-then-remove keeps every part payload;
  - chart reads on fixtures.
- `tests/xlsx_features.rs`: `auto_filter_column_criteria`.
- `tests/xlsb_remove.rs`: the removal capability list now includes the
  feature nodes.
- `crud_matrix`: XLSB claimed cells **58 → 175**, all ok.
- `corpus_fidelity`: the XLSB "note add+remove" structural round trip is
  value-exact on 18 fixtures and byte-exact on 1.
- Goldens re-blessed in both feature sets: capabilities and registry
  (`xlsb.json`, `xlsx*.json`), `officecli-capability-delta.md`, and
  `results/0289/{crud-checklist,fidelity}.{md,tsv}`.
- `pyo3-litchi`: `test_xlsb_worksheet_features`.

## 5. Measurements

Setup:
- Bench: `examples/xlsb_features_bench.rs`, release build, `taskset -c 7`.
- Three processes. Each runs leaf and unified ABBA and prints medians.
- Workbook: 40 notes, 40 conditional formats, 40 validations, 40 hyperlinks
  and 40 shapes.
- Raw lines: `.coord/xlsb-features/bench3.jsonl`, not committed.

| scenario | leaf µs | unified µs | ratios (3 runs) |
|---|---:|---:|---|
| `cf_get` | 48.8 | 51.2 | 1.048, 1.042, 1.044 |
| `comment_get` | 5.9 | 7.2 | 1.207, 1.206, 1.216 |
| `comment_add` | 573.7 | 574.7 | 1.000, 1.002, 1.002 |
| `cf_add` | 648.6 | 655.7 | 1.012, 1.013, 1.011 |
| `link_add` | 492.4 | 496.3 | 1.009, 1.007, 1.008 |
| `validation_add` | 590.1 | 596.3 | 1.004, 1.011, 1.006 |
| `shape_add` | 903.7 | 911.2 | 0.998, 1.005, 1.008 |
| `rowbreak_add` | 487.2 | 488.4 | 1.001, 1.008, 1.002 |

- Every write is within 1.3% of the leaf.
- Before this table, `rowbreak_add` was 1.096: each write ran the full
  feature pass, formula rendering included. `FeatureSet` fixed it (now
  1.002).
- `comment_get` is the one miss: +1.3 µs absolute, +21% relative. The
  adapter numbers notes row-major and builds the node's label. A lazy
  sheet-name lookup and an inline label already cut it down.

## 6. Gaps and follow-ups

1. XLSB chart writes. Charts are read only.
2. XLSB auto-filter criteria (`criteriaN.op`), x14 conditional formats and
   DV14 validations: authoring stays a gap. The existing records are kept.
3. The effect properties that XLSX also lacks (the 0320 gaps).
4. The XLSB `drawing::edit` duplicates the XLSX drawing editor, and its VML
   note helpers duplicate the XLSX ones. Follow-up: share them through a
   spreadsheet drawing crate.
5. An author added for a note stays in the comments author list after the
   note is removed.
6. DV formulas that reference another sheet fail to compile: the compiler
   runs without a workbook context.
7. The XLSB sheet `autoFilter` and cell `link` shortcut properties are still
   gaps. The element nodes cover both.
8. XLSX threaded comments are read in the leaf only. There is no unified
   node yet.
9. XLSX note font and rich-run formatting (`font.*`, `runs=`; brief item 3c)
   is not done. The XLSX comments splice writes plain text runs.
10. `comment_get` is +21% (§5).

## 7. Gates (after merging trunk `db96990ff6`)

Each command exited 0 unless noted. The build env was
`CARGO_BUILD_JOBS=6` and `RUST_TEST_THREADS=8`.

| command | exit | notes |
|---|---:|---|
| `cargo fmt --all --check` | 0 | |
| `cargo clippy -p litchi-xlsb -p litchi-xlsx -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 | |
| `cargo clippy -p litchi-ops --no-default-features -- -D warnings` | 0 | also with `--features xlsb` and `--features xlsx`, both 0 |
| `cargo test -p litchi-xlsb --all-features` | 0 | 1295 passed |
| `cargo test -p litchi-xlsx --all-features` | 0 | 2463 passed |
| `cargo test -p litchi-ops --all-features` | 0 | lib plus 144 test targets, in six chunks with executables deleted after each |
| `cargo test -p litchi --all-features` | 0 | 230 passed |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-xlsb -p litchi-xlsx -p litchi-ops` | 0 | |
| `python3 tools/check_crate_boundaries.py` | 0 | |
| `python3 tools/non_iwork_gate.py verify` | 0 | |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` | 0 | |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 | 282 tests |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 | |
| `cargo clippy -p litchi-ods --lib --all-features -- -D warnings` | 0 | lib only; its test suite needs a purged fixture |
| `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test officecli_corpus` | 0 | 5 passed; no outcome changed |
| `python3 tools/officecli_corpus.py`, then `--check` | 0, 0 | no change |
| `python3 tools/officecli_capability_delta.py`, then `--check` | 0, 0 | regenerated |

Two `litchi-ops` chunks failed on the first run and passed after a fix:

- `spec_tables_agree_with_their_vocabularies`: the XLSB axis spec lacked
  the `role` identity.
- `xlsb_capabilities_advertise_removal_for_cells_and_sheets`: the
  expectation predated the feature nodes.

Both were fixed in `f01d31bd6c` and pass on rerun (1202 passed).
