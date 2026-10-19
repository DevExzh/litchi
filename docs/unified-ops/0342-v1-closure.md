# 0342 — v1 corpus closure: DOCX table borders, row properties, list styles on add and application metadata; PPTX table-level set

- Status: implemented on `uo2/v1-closure`, cut from `feat/unified-ops-api`
  at `4282e69000`. Trunk was merged at `ed832736c3` (record 0336, the raw
  layer and `dump`) and again at `7ea8acee2a` (record 0341, the ADR 0008
  checklist rows); goldens, the CRUD checklist and its ADR 0008 rows,
  fidelity, the delta table and the corpus were re-blessed after each.
- Date: 2026-10-17
- Input: the v1-scope subset of the OfficeCLI corpus (record 0332;
  `officecli-corpus-deltas.md`), red with 68 open lines on the base tree:
  57 planned WP5 (record 0336) and the 11 this record closes. Records 0335
  F8 and 0339 §4 (WP4), 0253 F3 and 0338 §8 (DOCX application
  properties), and 0333 F5 (PPTX table-level `set`).
- Owned:
  - `crates/litchi-docx` (metadata family, table and row property splices,
    `Height::new`, `Refusal::UnmodelledProperty`);
  - `crates/litchi-pptx` (`opened::Edit::set_table_properties`);
  - `crates/litchi-ops/src/format/{docx*,pptx*}`;
  - the matching registry, checklist and fidelity rows;
  - `tools/officecli_dispositions.json`.
- Small hunks outside them: `format/metadata.rs` (DOCX core/application
  split of the shared admission), `capabilities.rs` (DOCX `tbl`/`tr` set
  claim), `transaction.rs` (the DOCX metadata step's change test also
  compares the application and custom parts), `error.rs` (the new
  `metadata::Error::Properties`).
- `performance_claim`: measured (§6). One path misses the 10% single-op
  budget (row set, 1.13), explained in §6.

## 1. The 11 lines

| Line | Was | Now |
|---|---|---|
| `word/tables.sh:189`–`194` | `set /body/tbl[4] border.{top,bottom,left,right,horizontal,vertical}`: `unsupported_element` | ok |
| `word/tables.sh:197` | `set …/tr[1] header=true height.exact=400`: `unsupported_element` | ok |
| `word/paragraph-formatting.sh:153` | `add paragraph listStyle=bullet`: warned (`listStyle`) | ok |
| `word/paragraph-formatting.sh:154` | `add paragraph listStyle=ordered start=5`: warned | ok |
| `word/document-formatting.sh:28` | `set / extended.{company,manager,template}`: warned | ok |
| `ppt/tables/tables-nested.sh:86` | `set /slide[2]/table[1] name firstRow … bandedCols`: `unsupported_capability` after skipping the switches | ok |

v1-scope subset (`python3 tools/officecli_corpus.py`):

| Tree | passed | accounted | open |
|---|---:|---:|---:|
| base `4282e69000` | 2,204 | 1,171 | 68 (57 WP5, 9 WP4, 1 0253, 1 0333) |
| this branch before the trunk merge | 2,217 | 1,169 | **57, all WP5** |
| trunk `ed832736c3` alone (record 0336) | 2,261 | 1,171 | 11 (the lines above) |
| this branch after the trunk merges (0336, 0341) | 2,274 | 1,169 | **0 — green** |

Two accounted lines moved to passed besides the 11:
`word/tables.sh:247` and `:272` (`add table border.all=…`, which skipped
`border.all` before). `word/tables.sh:179` still skips `cellSpacing`,
`indent`, `layout` and `padding` (deferred-P1), and
`ppt/tables/tables-nested.sh:54` still skips the table `style`
(deferred-P1, record 0333 F4).

Dispositions deleted or narrowed: `docx-app-metadata` and
`pptx-table-set` deleted; `docx-p0-residue` loses
`delta:docx:paragraph.listStyle`, `delta:docx:paragraph.start` and
`delta:docx:table.style`; `docx-structure-writes` loses the table and row
`unsupported_element` keys (cell insertion stays planned WP4). Both
`--check` runs are green.

## 2. DOCX application and custom properties

OfficeCLI's `ExtendedPropertiesHandler` writes `extended.company`,
`extended.manager` and `extended.template` to `docProps/app.xml`, and
`custom.<name>` to `docProps/custom.xml`.

- **Leaf.** `litchi_docx::metadata::Edit` stages, besides the core value,
  the application properties (`app_props`, `set_app_props`) and the custom
  properties (`custom_props`, `set_custom_props`), the xlsx spellings.
  `commit` computes the shared engine's exact part patch
  (`litchi_ooxml_common::properties::Patch`) for those two families, which
  splices only the changed scalar elements of `app.xml`, so everything the
  typed model does not carry keeps its bytes. `Patch` carries it
  (`Patch::parts`); `is_noop`, `inverse` and `apply_metadata` cover it. A
  value equal to the stored one publishes nothing; a change to a signed
  package is `Error::Signed`; an invalid application or custom graph is
  the new `Error::Properties`. `Package::app_props` reads the part.
- **Unified.** A root or `/metadata` set admits the full key set (the 22
  OfficeCLI spellings and `custom.<name>`, record 0254). Core keys go to
  the core value; `company`, `manager`, `template`, `application` and their
  `extended.*` spellings are merged onto the stored application
  properties; `contentType`, which no OOXML part stores, is
  `unsupported_property`.
- **Reads.** The DOCX root lists OfficeCLI's eleven `extended.*` keys
  (template, manager, company, application, applicationVersion, pages,
  words, characters, lines, paragraphs, totalTime). The three writable
  ones are settable; the others are read only. A source owner whose
  topology was edited reads none of them (its physical archive no longer
  answers for the part); `File::metadata()` stays the core properties.
- **Transactions.** A metadata step keeps an empty main-document patch
  (record 0338); the change test now also compares the application- and
  custom-part bytes of the endpoints, so an application-only step is a
  change, and `inverse` restores the input bytes
  (`tests/root_metadata.rs`).

## 3. DOCX tables and rows

- **Leaf.** `document::Edit::format_table_at(path, |&mut Table|)` and
  `format_row_at(path, |&mut Row|)` (owned and source-backed edits, and the
  snapshot edit) follow `format_paragraph_at`: parse the element's
  properties, apply the change, splice only the property elements whose
  typed value changed, read back, and record an invertible element
  replacement. The table edit writes `w:tblStyle` and `w:tblBorders` (the
  eight sides, `start`/`end` beside `left`/`right`); the row edit writes
  `w:trHeight` (value and rule), `w:cantSplit`, `w:tblHeader` and
  `w:hidden`, and a new `w:trPr` follows a row's `w:tblPrEx`. A change of
  any other property is `Refusal::UnmodelledProperty` (never a silent
  drop). `Height::new` constructs the non-exhaustive row height.
- **Unified.** Table and row nodes take `set` (they were
  `unsupported_element`):
  - table `border.all|top|bottom|left|right|horizontal|vertical`
    (write-only, set and add, `null` removes a side) with OfficeCLI's
    `ParseBorderValue` forms (`STYLE[;SIZE[;COLOR[;SPACE]]]` and `SIZE STYLE
    COLOR`; a bare size is eighths of a point, `pt`/`cm`/`in` convert; a
    missing size is 4, 12 for `thick`, none for `nil`; the style aliases
    `thin`, `medium`, `wavy`, …), and `style` (now set as well as add);
  - row `height` (no rule), write-only `height.exact` and `height.atLeast`,
    and `header`, `cantSplit`, `hidden` (`false` removes the flag, as
    OfficeCLI does). Bare heights are twips; units convert.
  - Every assignment of one table or row target is staged in one property
    edit (§6).
- **Add.** `border.*` seed a new table: the borders are written right after
  the insertion, in the same transaction.

## 4. DOCX `listStyle` and `start` on `add paragraph`

`listStyle` and `start` were set-only (record 0251e). They now seed a new
paragraph in the body, a table cell or a content control, after
OfficeCLI's add path (`FindContinuationNumId`): the new paragraph continues
the list of the nearest preceding paragraph of its container at the same or
a shallower level (deeper items are skipped) when that list has the same
kind; otherwise, and always with `start`, a new list definition is staged
(`Edit::add_list`). `none` adds a plain paragraph; the paragraph's own
`numLevel` is kept. The leaf already had the APIs (`add_list`,
`list_kind`, `format_paragraph_at`); no leaf change was needed. Headers,
footers and notes refuse a list style on `add`, as on `set`.

## 5. PPTX table-level `set`

- **Leaf.** `opened::Edit::set_table_properties(slide, shape,
  |&mut table::Properties|)` rewrites the `a:tblPr` switch attributes
  (`firstRow`, `firstCol`, `lastRow`, `lastCol`, `bandRow`, `bandCol`)
  lexically: a switch turned on is `"1"`, one turned off loses its
  attribute (off is the default), and a missing `a:tblPr` is created. Equal
  switches are an exact no-op; a change of the table style id is refused
  (record 0333 F4).
- **Unified.** `set /slide[N]/table[M]` takes `name` (the frame's
  `p:cNvPr@name`, through `set_shape_name`) and `firstRow`, `lastRow`,
  `firstCol`, `lastCol`, `bandedRows`, `bandedCols`. Other keys are
  `unsupported_property`, so `Warn` skips them (the table `style` and
  `border.*`, deferred-P1). Source owners are promoted, as row writes are.

## 6. Measurements

`examples/v1_closure_bench.rs`, release, `taskset -c 7`, eager handles,
ABBA, the median of 41 rounds, outputs compared outside timing; the leaf
leg is the hand-written edit doing the same work from the same snapshot.
Final file: [results/0342/v1-closure-bench.txt](results/0342/v1-closure-bench.txt)
(binary SHA-256 prefix in its header; run 2 agrees within 2%).

| Scenario | Leaf | Unified | Ratio |
|---|---:|---:|---:|
| DOCX `set / extended.company` | 37.1 µs | 37.3 µs | 1.01 |
| DOCX `set / title` (existing path, now split core/application) | 12.5 µs | 12.6 µs | 1.01 |
| DOCX table `set border.top` | 99.4 µs | 100.7 µs | 1.01 |
| DOCX row `set header height.exact` | 50.8 µs | 57.3 µs | **1.13** |
| DOCX `add paragraph listStyle=bullet` | 79.9 µs | 81.1 µs | 1.02 |
| PPTX table `set` (three switches) | 313.2 µs | 314.9 µs | 1.01 |

- The first row-set version was 1.62: it staged each assignment as its own
  `format_row_at`. All assignments of a table or row target now go through
  one property edit.
- **Miss.** The remaining 6.5 µs is path resolution: resolving
  `/body/tbl[1]/tr[1]` materializes the table's rows in the read model
  (4.4 µs, against 0.3 µs for `/body/tbl[1]`, the bench's attribution
  leg), and the element path proof then walks them. Follow-up 0342-P1:
  resolve a row path from the table's row index without building the row
  values.
- The DOCX `set / title` ratio stays where record 0338 measured it (1.03).
- Existing outputs are byte-identical: a core-only metadata set writes the
  same core part as before (the application patch is a no-op and is not
  applied), and the corpus fidelity run's byte-exact rows are unchanged
  (one row's set-then-revert probe now picks the root `extended.template`
  and reports the shared engine's strict refusal of an orphaned core part).

## 7. Deltas against OfficeCLI

- **V1.** A row flag set to `false` and a table switch set to `false` are
  removed, not written `"0"` (equivalent; fewer bytes).
- **V2.** A PPTX table `style` and table-level `border.*` are not written
  (deferred-P1, record 0333 F4).
- **V3.** A DOCX `border.*` value with an unknown style token is
  `invalid_value`; OfficeCLI passes it through to the XML.
- **V4.** A DOCX application-property write on a package whose
  property graph the shared engine refuses (an orphaned core part, for
  example) is refused; OfficeCLI writes.

## 8. Not claimed, and follow-ups

- **0342-P1.** Row path resolution builds the table's rows (§6).
- Table properties other than style and borders (`width`, `align`,
  `indent`, `cellSpacing`, `layout`, `padding`, `tblLook`, `caption`,
  `description`) and row properties other than height and the three flags
  are not written; they stay `unsupported_property` (deferred-P1).
- DOCX cell insertion (`add cell`), table columns, cell text keys and
  hyperlink keys stay planned WP4 (`docx-p0-residue`,
  `docx-structure-writes`).
- Reading `extended.*` from a source owner with edited topology.

## 9. Verification

On the merged tree; cargo through the memory-capped shim,
`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`. The commands and exit codes
are in this branch's final report: `fmt`; `clippy` (all features, and
`litchi-ops` without features and with `docx`, `pptx`); the
`litchi-docx`, `litchi-pptx`, `litchi-ops` (lib, doc, integration in
chunks) and `litchi` tests; the corpus test; `rustdoc`; the crate-boundary,
non-iWork and glossary gates; `pyo3`; `perf-baseline` `check --locked`;
the tool unit tests; both dispositions `--check` runs; the bench matrix;
and the feature matrix. Goldens were re-blessed in both feature sets
(all features, and every feature but `encryption`), and the CRUD checklist
and fidelity reports re-blessed: the DOCX table and row and the PPTX table
gain `set`, `edit` and `batch` cells.
