# 0295 — M4 engines on more formats, DOCX SmartArt, XLSX rich-text find

Status: implemented on `uo2/m4-ext` (cut from `feat/unified-ops-api` at
99cc3138eb, trunk 3dccbac8e8 merged in); awaiting coordinator review and
merge.

Inputs: [0273](0273-find-replace.md) (find/replace gaps),
[0274](0274-template-merge.md) (F1: merge on more formats; F4: XLSX
rich-text cells), [0279](0279-field-refresh.md) (F1: refresh on DOC, PPT,
RTF and ODF; F8: XLSB/XLS cache write-back), [0266](0266-csv-import-and-formula-evaluation.md)
(F8: ODS import), [0284](0284-typed-smartart-crud.md) (0284a: DOCX
SmartArt) and [0285](0285-find-merge-consolidation.md) (XLSX rich-text
gap). Glossary: [API_GLOSSARY.md](../API_GLOSSARY.md), section
"Find/replace and merge on every text format (record 0295)".

## 1. What landed

### One engine, per-format adapters (`litchi-ops/src/find/surface.rs`)

Find/replace and template merge for the formats whose leaf rewrites
whole runs, paragraphs or cells share one crate-private text surface: an
adapter lists the searchable text of its scope as `Run`s (text, group,
scoped, editable) in document order; `surface::find` plans the matches
with 0273's planner (leftmost, non-overlapping, a cross-run match
collapsing into the run where it starts, whole-word context across
unscoped runs) and `surface::merge` plans placeholders with 0274's planner
and 0285's range rewrite; the adapter stages one native replacement per
changed run in one native edit. A run the leaf cannot rewrite as text
(markup inside it) is `editable: false`: a find that touches it is refused
(`unsupported_capability`, so a broad find never silently leaves text
behind), and merge treats it as fixed content. Format properties on a find
are refused (`unsupported_property`) wherever the leaf cannot split a run.

| Format | Scope (find) | Unit (one group) | Native write | Merge |
|---|---|---|---|---|
| DOC | `/`, `/body`, a body paragraph, a run (run scope) | a main-story paragraph's character runs | `replace_text(TextTarget::BodySpan)` per run, last first, one edit | runs; values without line breaks; fields and drawings fixed; table text read and refused when a find would change it |
| RTF | as 0273 (now through the surface) | a body paragraph's runs | `replace_body_run_texts` | new: runs, values without line breaks |
| PPT | `/` (slides and notes), a slide, a shape, a paragraph, a notes page | one paragraph (between `\r` marks) | one `ShapeTextReplacement` per changed shape, `set_notes_text` per notes page | paragraphs; the native closure keeps a single-run body's formatting and refuses a length change where per-character ranges would shift |
| XLSB | workbook, sheet, cell, range, row, column | one stored text cell | `set_scalar(Text)` per cell, one cell-value edit per sheet, chained, one commit | text cells; an inline rich string or a formatted shared string is fixed |
| XLS | as XLSB | one stored, non-formula text cell | `set_value(Text)`, native edits of 1,024 cells chained in one mutation (as CSV import) | text cells; a cell showing a formatted shared string loses its runs (the leaf refuses dropping the last reference) |
| ODS | as XLSB, merged ranges too | one logical text cell (a repeated run is one cell per coordinate) | `set_cells` with the cell's style, merge role and range source kept | text cells; a cell holding hyperlinks is fixed |
| ODT, OTH, ODM | `/`, `/body` (comments excluded, as OfficeCLI's root), any body node by its byte range | a paragraph's or heading's runs; hyperlink edges and notes separate groups; bookmarks join | `Edit::set_text` per run (spaces, tabs and breaks re-encoded) | runs; tabs and line feeds become `text:tab`/`text:line-break` |
| ODP | `/` (slides and notes pages), a slide, an object, table, row, cell, paragraph, run, notes | a paragraph's runs; a link is its own group; fields are boundaries, never searched | `Change::SetRunText` per run, one checked splice | runs |

`File::set` with `find=`, `transaction::Edit::set`, batch `set` items and
`File::merge`/`Edit::merge` dispatch to these adapters; the XLS ones commit
several native steps as one mutation (`Edit::execute_steps`, and
`run_steps` in the merge and refresh transactions). Exact no-ops (no
match, no placeholder) make no change; every commit keeps its inverse.

### XLSX rich-text cells (0285 gap, 0274 F4)

`litchi-xlsx` gains `RichText`/`RichRun`, `Worksheet::rich_text` (the runs
of a formatted shared-string item, or of an inline rich string read from
the worksheet part, each run's `rPr` children kept as stored; `None` for
phonetic runs or nested property markup) and
`WorksheetEdit::set_rich_text` (a new `Payload::Rich`: the runs written as
the cell's own inline string, `rPr` children re-emitted with the cell's
namespace prefix). XLSX find rewrites such a cell run by run instead of
refusing it; XLSX merge fills its runs instead of flattening the cell.

### CSV import into ODS (0266 F8)

`File::import` writes ODS through one leaf edit (`set_cells` in batches):
numbers as `float`, booleans as `boolean`, ISO dates as native `date`
cells (`office:date-value`, no serial or number format needed), text as
`string`; a cell's style and merge role are kept, an empty field clears a
stored cell. `formulas(true)` is refused (OpenFormula translation is a
follow-up); ODS evaluation stays a gap with that follow-up.

### Field refresh (0279 F1, F8)

| Format | What is computed | How it is written |
|---|---|---|
| DOC | the main story's top-level fields with a result, by the Word field computation now shared with DOCX (`refresh::word_field`: `DATE`/`TIME` with a clock, `CREATEDATE`/`SAVEDATE`/`PRINTDATE`, the document-property fields, `NUMWORDS`/`NUMCHARS` with field instructions left out, `SEQ`); `PAGE`-like fields `needs_layout`; `REF` `unsupported` | `replace_text(TextTarget::FieldResult)`, one edit; a story with a nested field or a field without a result reports every computable field `unsupported` (the native edit cannot address them) |
| ODT, OTH, ODM | `text:date`/`text:time` (clock), stored-time and document-property fields (`title`, `subject`, `description`, `keywords`, `initial-creator`, `creator`); page fields `needs_layout`; `text:fixed` fields `locked`; fields with an ODF data style `unsupported` | new leaf `body::text_fields` and `Edit::set_field_text`: the field's content replaced, element and attributes kept, as a checked splice |
| ODP | as ODT, plus `text:page-number` (the slide position, `text:select-page` and `style:num-format` `i`/`I`/`a`/`A` applied) and `text:page-count` (the slide count) on every slide and notes page | new leaf `scene::Change::SetFieldText` |
| XLSB | formula results (0279's recalculation) written back with `write_cache` | the cell-value edit's formula caches; the four `BrtFmla*` families are now one family for a cache write (the record re-encoded with its formula) |
| XLS | as XLSB | `Value::FormulaCache` through chained native edits |

PPT and RTF refresh stay gaps (§5).

### DOCX SmartArt (0284a)

- Leaf `litchi-docx`: `document::NewSmartArt` and
  `Edit::insert_smartart_at(paragraph, at, &NewSmartArt)` stage the five
  diagram parts from the authored `Spec` of record 0284 (data model with
  its drawing reference, Litchi's layout definition, quick style, colors,
  cached drawing), their five main-document relationships (transitional
  or Strict types) and the inline `dgm:relIds` drawing run, in one graph
  change and one document operation; the result is the data relationship
  id `Package::smart_arts` reports.
- Unified: a WORD-vocabulary `smartart` element addressed document-wide
  (`/smartart[N]`, `/body/smartart[N]`), listed after the comments part
  among the root's children, read from the leaf inventory (`layout`,
  `layoutId`, `nodeCount`, `text` as the outline) for eager and
  source-backed owners; `add` under a paragraph (`layout`, `text`, `name`,
  `alt`, `width`, `height`). `set` and `remove` are gaps (§5). OfficeCLI's
  Word `diagram` is its mermaid renderer, so `diagram` is not an alias.

### Leaf hunks

| Crate | Change |
|---|---|
| `litchi-xlsx` | `RichText`, `RichRun`, `Worksheet::rich_text`, `WorksheetEdit::set_rich_text`, `Payload::Rich` and its writer |
| `litchi-xlsb` | formula caches may change type in a cell-value edit |
| `litchi-odt` | `body::{TextField, text_fields}`, `Edit::set_field_text` |
| `litchi-odp` | `scene::Change::SetFieldText` |
| `litchi-docx` | `NewSmartArt`, `Edit::insert_smartart_at`, `next_relationship_ids` |
| `litchi-odf-common` | `XmlSplicePublication::replace` admits an edit after every earlier one in O(1) (it scanned every earlier edit: quadratic in the number of splices) |

### Capabilities, schema, surfaces

- `replace` (find) and `merge` are supported for DOC, PPT, XLS, XLSB, RTF,
  ODT (and OTH/ODM), ODS and ODP; `refresh` for DOC, ODT and ODP, and
  XLSB/XLS now write caches; `import` for ODS. The RTF find capability was
  reported as a gap since record 0273 although the adapter worked: fixed.
- Goldens re-blessed in both sets (`--all-features` and every feature but
  `encryption`); the OfficeCLI capability delta table regenerated; the
  OfficeCLI corpus re-recorded with no outcome change (only the
  classification text of the `word/diagram.sh` rows moves).
- CRUD checklist: a `smartart` recipe for DOCX; the docx row gains the
  element (7 claimed cells, all `ok`).
- CLI and Python: the commands and methods dispatch through the engines
  unchanged; help texts, docstrings and stubs name the new formats; tests
  cover XLSB/ODT/ODS merge, ODT find, ODS import and the new refusals.

## 2. Tests

- `tests/find_merge_formats.rs` (13): DOC (runs keep bold, scopes, no-op,
  merge across runs, `Missing::Empty`), RTF merge, PPT shapes and notes,
  XLSB/XLS/ODS (workbook, sheet and cell scopes, numbers untouched, merge on
  two sheets), XLS merge of 1,500 cells across chained edits, ODT/OTH/ODM
  runs (bold kept, links, notes searched, comments not, run scope, merge
  with line breaks), ODP slides and paragraph scopes.
- `tests/xlsx_find_replace.rs`: rich shared strings and inline rich strings
  rewritten run by run (`rPr` bytes checked in the worksheet XML), a
  cross-run match, merge keeping runs.
- `tests/field_refresh_formats.rs`: DOC (statuses, one commit, inverse,
  convergence to an exact no-op), ODT (`variable.odt`: locked fixed fields,
  title from the metadata, page fields), ODP (page numbers by position,
  after a slide swap).
- `tests/xlsx_refresh.rs`: XLSB and XLS cache write-back, once, inverse.
- `tests/ods_import.rs`, `tests/docx_smartart.rs` (eager and source-backed
  owners, reopen, refusals, capabilities).
- Leaf units: `surface` engine, XLSX `rich_item` and the rich write,
  XLSB cache type change with exact inverse, ODT text fields, ODP
  `SetFieldText`, DOCX `insert_smartart_at` doctest, ODS date lexicals.

## 3. Performance

`m4_formats_bench` (new; release, `taskset -c 7`, two warm-up rounds then 9
ABBA rounds, medians, two runs; load average 5.4–5.6):
[results/0295/](results/0295/) (all samples; binary SHA-256 in
`binaries.sha256`, built from d65f5f9608).

| Scenario | Size | Leaf | Unified | Ratio (run 1 / run 2) |
|---|---:|---:|---:|---|
| ODT find, 1 match | 1,000 ¶ | 418 µs | 453 µs | 1.082 / 1.084 |
| ODT find, every ¶ | 1,000 ¶ | 1.06 ms | 1.07 ms | 1.009 / 1.003 |
| ODT merge | 1,000 ¶ | 0.98 ms | 1.00 ms | 1.020 / 1.014 |
| ODT find, 1 match | 10,000 ¶ | 2.51 ms | 2.81 ms | 1.117 / 1.121 |
| ODT find, every ¶ | 10,000 ¶ | 8.79 ms | 8.74 ms | 0.991 / 0.997 |
| ODT merge | 10,000 ¶ | 7.91 ms | 8.18 ms | 1.036 / 1.032 |
| DOC find, 1 match | 1,000 ¶ | 1.99 ms | 2.02 ms | 1.018 / 1.011 |
| DOC find, 100 matches | 1,000 ¶ | 51.6 ms | 52.2 ms | 0.998 / 1.024 |
| DOC find, 1 match | 5,000 ¶ | 24.0–30.3 ms | 23.9–30.4 ms | 0.999 / 1.002 |
| DOC find, 100 matches | 5,000 ¶ | 370–377 ms | 334–338 ms | 0.902 / 0.896 |
| ODS import | 1,000 rows × 4 | 23.0 ms | 23.4 ms | 1.021 / 1.013 |
| ODS import | 10,000 rows × 4 | 242–251 ms | 256–269 ms | 1.058 / 1.072 |

Leaf legs: ODT walks the body index and rewrites each character-data run
holding the needle (one `Edit`); DOC scans the reader's body paragraph
texts and replaces whole paragraphs (a lower bound: one-run paragraphs);
ODS splits the CSV by hand and stages typed cells with `set_cells`.

- The ODT and ODP splice publication was quadratic in the number of
  splices: 10,000 run rewrites took 20.7 ms, of which most was
  `XmlSplicePublication::replace`'s overlap scan (callgrind). With the
  O(1) check for ascending edits both legs take 8.6–8.8 ms (2.4× faster);
  the change is in the shared leaf, so every ODF splice writer benefits.
- DOC find first measured 1.4–1.9× against the leaf. Callgrind: the
  unified leg read every paragraph's runs, and the `File::apply` path pays
  the transaction admission's probe edit (`transaction::admit` builds and
  drops a native DOC edit: 18% of the profile). Reading runs only of
  paragraphs whose visible text may match brought `File::set` to 1.00–1.02;
  the admission probe is shared by every DOC transaction and is left as a
  follow-up (F6). The 100-match DOC case is faster unified because the
  adapter rewrites only the matched run, the leaf leg whole paragraphs.
- ODT find with one match in 10,000 paragraphs is 11–12% over the leaf
  leg, above the 10% single-operation budget (a find is a scan, not a
  single set). The excess is the surface's per-run bookkeeping (run, key and
  piece vectors and the group planner) over a scan the leaf leg does with
  `str::contains`; follow-up F7. Every other scenario is within 7%.
- ODS import of 40,000 cells is 6–7% over the leaf: parsing with type
  inference and the 16 MiB staging checks the hand split omits.
- Not measured: XLSB/XLS/PPT/ODP find and merge, XLSX rich-text find, the
  refresh adapters and DOCX SmartArt insertion (they reuse the measured
  write paths of records 0281/0286/0283; follow-up F8).

## 4. Gates

All with `CARGO_TARGET_DIR=<worktree>/target`, the build environment of
the rules (`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`), one cargo command
at a time; exit 0 unless noted.

- `cargo fmt --all --check`
- `cargo clippy -p litchi-docx -p litchi-xlsx -p litchi-xlsb -p litchi-odt
  -p litchi-odp -p litchi-cli -p litchi-ops -p litchi -p litchi-py
  --all-targets --all-features -- -D warnings` (first run: one
  `clone_on_copy`, fixed)
- `cargo clippy -p litchi-ops --no-default-features [--features f]
  --all-targets -- -D warnings` for no feature and each of `doc`, `ppt`,
  `xls`, `xlsb`, `rtf`, `odt`, `ods`, `odp`, `oth`, `odm`, `docx`, `xlsx`,
  `pptx`, `xlsb,eval`, `xls,eval`, `doc,ppt,xls,rtf` (the first
  no-feature run failed on two ungated test files and the `doc`, `rtf` and
  `ppt` runs on an unused test import; fixed and rerun), and `cargo clippy
  -p litchi-ods --lib -- -D warnings`
- tests: `cargo test -p litchi-ops --all-features` (lib 905, doc 204,
  integration 1,144, run in four chunks with the test binaries pruned
  between them), `-p litchi`
  (260), `-p litchi-docx` (2,394), `-p litchi-xlsx` (2,449), `-p
  litchi-xlsb` (1,273), `-p litchi-odt` (1,114), `-p litchi-odp` (435),
  `-p litchi-cli` (one CLI test expected XLSB merge to fail, updated),
  `-p litchi-odf-common` (573); `litchi-odt`, `litchi-odp` and
  `litchi-odf-common` rerun after the splice change
- `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features` for the
  changed crates
- `python3 tools/check_crate_boundaries.py`; `python3
  tools/non_iwork_gate.py verify`
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir
  target/api-glossary`
- `crates/pyo3-litchi/tests/run_tests.sh` (271)
- `cd tools/perf-baseline && cargo check --locked --all-targets`
- capability and registry goldens in both sets, `python3
  tools/officecli_capability_delta.py --officecli …`
- `LITCHI_BLESS=1 cargo test -p litchi-cli --all-features --test
  officecli_corpus`, `python3 tools/officecli_corpus.py`, `--check`
- `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test
  crud_matrix`

## 5. Gaps and follow-ups

- F1 refresh for PPT and RTF: no leaf edit addresses their field results
  (PPT slide-number atoms are not text; RTF has no field-result edit).
- F2 formatting on a find (`bold=true` with `find=`) for DOC, PPT, RTF,
  ODT, ODP and the workbooks: the leaves format whole runs and cannot split
  one at a match.
- F3 ODF headers and footers (`styles.xml`), ODT comments as a find scope,
  DOC headers, footers, notes and tables (refused when a find would change
  table text), PPT shapes inside groups, ODP master pages and layouts.
- F4 ODS cells keep no span formatting (the leaf cell model has none) and
  cells with hyperlinks are fixed; XLS cells showing a formatted shared
  string become plain text; ODS formulas need OpenFormula translation for
  import and evaluation.
- F5 DOCX SmartArt `set` and `remove` (part replacement and drawing removal
  for diagrams in the leaf), paragraph-local addresses
  (`/body/p[N]/smartart[M]`) and the frame's name, alternative text and
  extent in the read projection.
- F6 the DOC transaction admission builds and drops a native edit
  (`transaction::admit`); `File::apply`/`batch` on DOC pay it once more.
- F7 ODT find over many unmatched paragraphs: a paragraph prefilter (as
  DOC has) without a second text decode.
- F8 paired measurements for the XLSB, XLS, PPT and ODP adapters, XLSX
  rich-text rewrites, the refresh adapters and DOCX SmartArt insertion.
- ODT/ODP date fields with an ODF data style are not refreshed (data
  styles are not applied); `text:author-name` and the statistics fields
  are `unsupported`.

## 6. Not claimed

- Rendering of the refreshed ODF fields beyond their stored text, or that
  LibreOffice keeps a refreshed value of a non-fixed field (it recomputes
  them on load).
- That the unified find is faster than a hand-written leaf find; the
  measured overheads are above.
