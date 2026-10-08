# S8b — perf-baseline routes for `litchi::open` and `File`

- Status: done on `uo/m0a-s8b-perf`, branched from `uo/m0a` at `20692d6444`
  (S6, S6b and S7 merged).
- Scope: plan [0003](0003-m0a-migration-plan.md) §4 step S8b, §5.2 and §5.3;
  goal [`UNIFIED_OPS_GOAL.md`](../UNIFIED_OPS_GOAL.md), PERFORMANCE BAR.
- Inputs: the S1 baseline
  [`results/0003/before`](results/0003/before/README.md) (route list, selector
  names, noise findings) and record 0010 §9.1 (naming) and finding F-11.

No measurement is reported here. The smoke run below shows that each route
runs. Comparisons belong to the S10 record.

## What changed

`tools/perf-baseline` gains 118 opt-in selectors. For every §5.3 scenario
(open, list, read one, full text, detection) there are two legs:

- a **unified** leg that goes through `litchi::open*`, `litchi::File`, its
  family views and its format handles;
- a **leaf** leg that does the same work directly on the leaf crate.

The old routes stay until S9 deletes the old facade types. The harness's
predecessor code is unchanged, and so are their frozen records.

| File | Change |
|---|---|
| `tools/perf-baseline/src/unified_ops.rs` (new) | The route catalog: 118 entries, each with its leg, predecessor, twin and plan. Also the shared leaf helpers, the in-process runners (XLSX bytes, DOCX read-one, DOC/PPT, RTF, ODT, ODP, ODS, detection) and the per-result `RouteSummary`. |
| `tools/perf-baseline/src/filesystem/ops.rs` (new) | The 54 DOCX, PPTX and XLSX filesystem routes in the fresh-child protocol: prepared roots, timed bodies, untimed verification and the XLSX evidence. |
| `tools/perf-baseline/src/filesystem.rs` | Adds `Operation::Ops(ops::Route)`. The format predicates, cold-verified eligibility, selected-cell and timed-text evidence checks, source replays and output verification learn the new variant; every predecessor path is unchanged. |
| `tools/perf-baseline/src/lib.rs` | Adds the `Case` variants, `Case::name`, `parse_case`, `--help`, `is_unified_ops_{filesystem,in_process}`, the in-process dispatch, `SourceSummary::unified_ops`, and an exhaustive `run_case_with_config` arm. `EXPECTED_SELECTABLE_CASE_COUNT` goes from 534 to 652. |
| `tools/perf-baseline/tests/unified_ops_filesystem.rs` (new) | Runs the six XLSX file pairs and two XLSX byte pairs through the binary and checks the fresh children, the pinned source and semantic digests, the selected-cell digest and the pair metadata. |
| `tools/perf-baseline/Cargo.toml`, `Cargo.lock` | Adds a direct `litchi-ops` dependency for the detection leg. The lock had been stale since S4 made `litchi` depend on `litchi-ops`, so `--locked` builds failed; it now records `litchi-ops`. No version changes. |
| `tools/perf-baseline/src/producer_shape/selected_route_tests.rs` | Adds `#[allow(clippy::result_large_err)]` to a test helper. `-D warnings` clippy already failed there at the base commit, because the leaf scanner's error type has grown. |
| `docs/performance/crud-coverage-index-v2.json` | Binds 652 names in `Case::name` order with a new digest. The 119 new names are `excluded`, reason `not-selected-in-representative-matrix`. |
| `tools/test_crud_coverage_index.py` | Pins 652 and the new digest. This also fixes **F-11**: the index and test pinned 533 while `lib.rs` registered 534 (`xlsx_producer_medium_source_full_cell_scan`, change 0884, was never indexed). |
| `tools/test_perf_baseline_source_policy.py` | Adds two tests. One checks that the new modules hold no `unsafe`, `std::env`, `Command` or global allocator, and that they read caller-named files only through the bounded reader. The other checks that every new name is its predecessor's plus the leg infix, that each leg has its twin (except six detection routes), that the predecessors are still registered, and that the new routes stay out of `Case::DEFAULT`. |
| `tools/perf-baseline/README.md` | Adds the section "Paired unified-API and leaf-API routes (M0a step S8b)", including how S10 uses the pairs. |
| `docs/unified-ops/results/0003/s8b/route-map.json` (new) | The machine-readable form of the route table below, for the S10 driver. |
| `docs/unified-ops/results/0003/s8b/smoke-summary.json` (new) | Which routes ran in the smoke run, with the commands and binary identity. It holds no timings. |

## Naming

The name is the predecessor's with the leg after the format family prefix:
`<prefix>_unified_<rest>` and `<prefix>_leaf_<rest>`. For example,
`docx_file_source_full_text` becomes `docx_file_unified_source_full_text` and
`docx_file_leaf_source_full_text`. The 0638 OLE2 facade family drops
`facade`: `doc_facade_file_open` becomes `doc_file_unified_open` and
`doc_file_leaf_open`. Unlike plan 0003 §5.3's example
(`docx_file_unified_full_text`), the names keep the predecessor's
`eager`/`source` word, because both owner routes are re-measured. `eager`
keeps the predecessor's meaning:

- for DOCX, the owning `litchi_docx::Package` (`SourceMode::Eager` on the
  unified leg);
- for PPTX, ODT, ODP, ODS and XLSX bytes, the in-memory byte route.

Record 0010 §9.1 reserves the same infixes for the M1 routes.

## What each leg times

The **unified** leg opens with `litchi::open(path)`, `litchi::open_bytes(bytes)`
or `litchi::open_bytes_with(bytes, SourceMode::Eager)`, then reads through
`File::text`, `File::metadata`, `File::document()`, `File::presentation()`,
`File::workbook()`, `Xlsx::sheet` and `Pptx::slide_size`.

The **leaf** leg makes the leaf calls the unified route makes, directly. The
format is known, so it runs:

- no format detection;
- no signature or encryption classification;
- none of the unified layer's own source-freshness re-checks;
- no handle construction.

Leaf work the unified route does at open is included:

- the eager DOCX owner's one-pass text validation;
- the core-properties read the XLSX open makes;
- the OLE property-set read the DOC and PPT opens make.

A pair's difference is therefore the unified layer's own work: detection, the
handle, view dispatch, freshness checks and wrapper conversion. The leaf legs
skip the work the predecessor routes added on top of the leaf, such as
`detection_smart` in `docx_file_eager_*`, so they are not interchangeable with
those predecessors.

| Family | Unified leg | Leaf leg |
|---|---|---|
| DOCX eager | `fs::read` + `open_bytes_with(Eager)`; queries `document()?.paragraph_count()`, `.paragraphs()`, `File::text` | `fs::read` + `OpcPackage::from_bytes_with_limits` + `litchi_docx::Package::from_opc_package` + `document()?.text()`; queries on `Package::document()` |
| DOCX source | `litchi::open(path)`; the same queries | `source_backed::Package::open`; `document()?.{paragraph_count, paragraphs, extract_text}` |
| PPTX eager (bytes) | `fs::read` + `litchi::open_bytes`; `presentation()?.{slides, slide_count, slide(100)}` | `fs::read` + `SourceBackedPresentation::from_read_at(OwnedSource)`; `slides()`, `text_and_name`, `slide_count`, `slide(100)` |
| PPTX source | `litchi::open(path)`; the same queries | `SourceBackedPresentation::from_path`; the same leaf queries |
| XLSX file | `litchi::open(path)`; `workbook()?.{worksheet_names, worksheet_count}` + `File::text`; `Xlsx::sheet(mixed case)` + `SelectedWorksheet::cell` | `SourceBackedPackage::from_read_at_with_limits(FileSource)` + core-properties read + `SourceBackedWorkbook::from_source_backed_package`; names, `len`, a row-by-row cell-text projection; `sheet(mixed case)` + `cell` |
| XLSX bytes | `litchi::open_bytes`; the same lifecycle | the same leaf open over `OwnedSource` |
| DOCX read one | eager `File`: `document()?.paragraph(middle)`, or `Paragraph::text` on a prepared paragraph; producer: source-backed `File`, `document()?.paragraph_text(i)` | the same calls on the leaf owner the unified route builds |
| DOC | fresh `litchi::open(path)` + `File::text` or `document()?.paragraph_text(target)` | bounded read + `OleFile::open(Cursor)` + `litchi_doc::Package::from_ole_file` + `document()` + property sets; `text()`, `paragraphs().nth(target)` |
| PPT | fresh `litchi::open(path)` + `File::text` or `presentation()?.slide(target)` | `litchi_ppt::SourceBackedPackage::from_path` + property sets + `presentation()`; `text()`, `slide_at(target)` + `text` + `shape_count` |
| RTF | `litchi::open_bytes` (+ `File::text`) | `litchi_rtf::RtfDocument::from_owned_bytes` (+ `text`) |
| ODT | `fs::read` + `open_bytes`, or `open(path)` (+ `File::text`) | `litchi_odt::Document::from_bytes` or `SourceBackedDocument::from_path` (+ `text`) |
| ODP | `open_bytes(owned)` or `open(path)`; `presentation()?.slide(middle)` | `litchi_odp::Presentation::from_bytes` or `SourceBackedPresentation::from_path`; `slide(middle)` |
| ODS | `open_bytes(owned)` or `open(path)` | `litchi_ods::Spreadsheet::from_bytes` or `SourceBackedSpreadsheet::from_path` |
| Detection | `litchi_ops::detect::{detect_file_format_from_bytes, detect_format_from_reader}`, the coordinator `litchi::open` runs, without the facade's iWork tail | `litchi_odf_common::detect::{bytes, reader}` (ODF byte and reader inputs only) |

Ingress, untimed preparation and timer placement follow the predecessor:

- query routes prepare their root outside the timer;
- DOCX and PPTX roots drop inside the timer;
- XLSX roots stay alive until every operation-only snapshot is taken;
- DOC and PPT routes open the file afresh in every sample.

## Route table

The table lists all 62 predecessors. 56 are paired with a unified and a leaf
leg; the 6 detection polyglot and catalog-alias inputs have a unified leg only.
The last column gives the smoke-run result for the unified and leaf legs.

| Scenario | Predecessor (kept until S9) | Unified leg | Leaf leg | Smoke run: ran |
|---|---|---|---|---|
| open | `docx_file_eager_open` | `docx_file_unified_eager_open` | `docx_file_leaf_eager_open` | yes / yes |
| open | `docx_file_source_open` | `docx_file_unified_source_open` | `docx_file_leaf_source_open` | yes / yes |
| list | `docx_file_eager_paragraph_count` | `docx_file_unified_eager_paragraph_count` | `docx_file_leaf_eager_paragraph_count` | yes / yes |
| list | `docx_file_source_paragraph_count` | `docx_file_unified_source_paragraph_count` | `docx_file_leaf_source_paragraph_count` | yes / yes |
| list | `docx_file_eager_list_paragraphs` | `docx_file_unified_eager_list_paragraphs` | `docx_file_leaf_eager_list_paragraphs` | yes / yes |
| list | `docx_file_source_list_paragraphs` | `docx_file_unified_source_list_paragraphs` | `docx_file_leaf_source_list_paragraphs` | yes / yes |
| full text | `docx_file_eager_full_text` | `docx_file_unified_eager_full_text` | `docx_file_leaf_eager_full_text` | yes / yes |
| full text | `docx_file_source_full_text` | `docx_file_unified_source_full_text` | `docx_file_leaf_source_full_text` | yes / yes |
| list | `docx_file_eager_open_paragraph_count_lifecycle` | `docx_file_unified_eager_open_paragraph_count_lifecycle` | `docx_file_leaf_eager_open_paragraph_count_lifecycle` | yes / yes |
| list | `docx_file_source_open_paragraph_count_lifecycle` | `docx_file_unified_source_open_paragraph_count_lifecycle` | `docx_file_leaf_source_open_paragraph_count_lifecycle` | yes / yes |
| full text | `docx_file_eager_open_full_text_lifecycle` | `docx_file_unified_eager_open_full_text_lifecycle` | `docx_file_leaf_eager_open_full_text_lifecycle` | yes / yes |
| full text | `docx_file_source_open_full_text_lifecycle` | `docx_file_unified_source_open_full_text_lifecycle` | `docx_file_leaf_source_open_full_text_lifecycle` | yes / yes |
| open | `pptx_file_eager_open` | `pptx_file_unified_eager_open` | `pptx_file_leaf_eager_open` | yes / yes |
| open | `pptx_file_source_open` | `pptx_file_unified_source_open` | `pptx_file_leaf_source_open` | yes / yes |
| list | `pptx_file_eager_list_slides` | `pptx_file_unified_eager_list_slides` | `pptx_file_leaf_eager_list_slides` | yes / yes |
| list | `pptx_file_source_list_slides` | `pptx_file_unified_source_list_slides` | `pptx_file_leaf_source_list_slides` | yes / yes |
| list | `pptx_file_eager_slide_count` | `pptx_file_unified_eager_slide_count` | `pptx_file_leaf_eager_slide_count` | yes / yes |
| list | `pptx_file_source_slide_count` | `pptx_file_unified_source_slide_count` | `pptx_file_leaf_source_slide_count` | yes / yes |
| read one | `pptx_file_eager_selected_slide` | `pptx_file_unified_eager_selected_slide` | `pptx_file_leaf_eager_selected_slide` | yes / yes |
| read one | `pptx_file_source_selected_slide` | `pptx_file_unified_source_selected_slide` | `pptx_file_leaf_source_selected_slide` | yes / yes |
| list | `pptx_file_eager_open_slide_count_lifecycle` | `pptx_file_unified_eager_open_slide_count_lifecycle` | `pptx_file_leaf_eager_open_slide_count_lifecycle` | yes / yes |
| list | `pptx_file_source_open_slide_count_lifecycle` | `pptx_file_unified_source_open_slide_count_lifecycle` | `pptx_file_leaf_source_open_slide_count_lifecycle` | yes / yes |
| read one | `pptx_file_eager_open_selected_slide_lifecycle` | `pptx_file_unified_eager_open_selected_slide_lifecycle` | `pptx_file_leaf_eager_open_selected_slide_lifecycle` | yes / yes |
| read one | `pptx_file_source_open_selected_slide_lifecycle` | `pptx_file_unified_source_open_selected_slide_lifecycle` | `pptx_file_leaf_source_open_selected_slide_lifecycle` | yes / yes |
| open | `xlsx_file_open` | `xlsx_file_unified_open` | `xlsx_file_leaf_open` | yes / yes |
| open | `xlsx_file_open_lifecycle` | `xlsx_file_unified_open_lifecycle` | `xlsx_file_leaf_open_lifecycle` | yes / yes |
| read one | `xlsx_file_selected_cell` | `xlsx_file_unified_selected_cell` | `xlsx_file_leaf_selected_cell` | yes / yes |
| open | `xlsx_bytes_open` | `xlsx_bytes_unified_open` | `xlsx_bytes_leaf_open` | yes / yes |
| open | `xlsx_bytes_open_lifecycle` | `xlsx_bytes_unified_open_lifecycle` | `xlsx_bytes_leaf_open_lifecycle` | yes / yes |
| read one | `docx_semantic_one_paragraph` | `docx_semantic_unified_one_paragraph` | `docx_semantic_leaf_one_paragraph` | yes / yes |
| read one | `docx_semantic_one_paragraph_text` | `docx_semantic_unified_one_paragraph_text` | `docx_semantic_leaf_one_paragraph_text` | yes / yes |
| read one | `docx_producer_source_selected_paragraph` | `docx_producer_unified_source_selected_paragraph` | `docx_producer_leaf_source_selected_paragraph` | yes / yes |
| open | `doc_facade_file_open` | `doc_file_unified_open` | `doc_file_leaf_open` | yes / yes |
| full text | `doc_facade_file_full_text` | `doc_file_unified_full_text` | `doc_file_leaf_full_text` | yes / yes |
| read one | `doc_facade_file_one_paragraph` | `doc_file_unified_one_paragraph` | `doc_file_leaf_one_paragraph` | yes / yes |
| open | `ppt_facade_file_open` | `ppt_file_unified_open` | `ppt_file_leaf_open` | yes / yes |
| full text | `ppt_facade_file_full_text` | `ppt_file_unified_full_text` | `ppt_file_leaf_full_text` | yes / yes |
| read one | `ppt_facade_file_one_slide_text` | `ppt_file_unified_one_slide_text` | `ppt_file_leaf_one_slide_text` | yes / yes |
| open | `rtf_file_open` | `rtf_file_unified_open` | `rtf_file_leaf_open` | yes / yes |
| open | `rtf_file_open_lifecycle` | `rtf_file_unified_open_lifecycle` | `rtf_file_leaf_open_lifecycle` | yes / yes |
| open | `odt_file_eager_open` | `odt_file_unified_eager_open` | `odt_file_leaf_eager_open` | yes / yes |
| open | `odt_file_source_open` | `odt_file_unified_source_open` | `odt_file_leaf_source_open` | yes / yes |
| full text | `odt_file_eager_open_full_text_lifecycle` | `odt_file_unified_eager_open_full_text_lifecycle` | `odt_file_leaf_eager_open_full_text_lifecycle` | yes / yes |
| full text | `odt_file_source_open_full_text_lifecycle` | `odt_file_unified_source_open_full_text_lifecycle` | `odt_file_leaf_source_open_full_text_lifecycle` | yes / yes |
| open | `odp_file_eager_open` | `odp_file_unified_eager_open` | `odp_file_leaf_eager_open` | yes / yes |
| open | `odp_file_source_open` | `odp_file_unified_source_open` | `odp_file_leaf_source_open` | yes / yes |
| read one | `odp_file_eager_selected_slide` | `odp_file_unified_eager_selected_slide` | `odp_file_leaf_eager_selected_slide` | yes / yes |
| read one | `odp_file_source_selected_slide` | `odp_file_unified_source_selected_slide` | `odp_file_leaf_source_selected_slide` | yes / yes |
| open | `ods_file_eager_open` | `ods_file_unified_eager_open` | `ods_file_leaf_eager_open` | yes / yes |
| open | `ods_file_source_open` | `ods_file_unified_source_open` | `ods_file_leaf_source_open` | yes / yes |
| detection | `detect_odt_bytes` | `detect_odt_unified_bytes` | `detect_odt_leaf_bytes` | yes / yes |
| detection | `detect_odt_reader` | `detect_odt_unified_reader` | `detect_odt_leaf_reader` | yes / yes |
| detection | `detect_odt_polyglot` | `detect_odt_unified_polyglot` | — (decided by the coordinator's precedence; compare with the predecessor) | yes / — |
| detection | `detect_odt_catalog_alias` | `detect_odt_unified_catalog_alias` | — (decided by the coordinator's precedence; compare with the predecessor) | yes / — |
| detection | `detect_ods_bytes` | `detect_ods_unified_bytes` | `detect_ods_leaf_bytes` | yes / yes |
| detection | `detect_ods_reader` | `detect_ods_unified_reader` | `detect_ods_leaf_reader` | yes / yes |
| detection | `detect_ods_polyglot` | `detect_ods_unified_polyglot` | — (decided by the coordinator's precedence; compare with the predecessor) | yes / — |
| detection | `detect_ods_catalog_alias` | `detect_ods_unified_catalog_alias` | — (decided by the coordinator's precedence; compare with the predecessor) | yes / — |
| detection | `detect_odp_bytes` | `detect_odp_unified_bytes` | `detect_odp_leaf_bytes` | yes / yes |
| detection | `detect_odp_reader` | `detect_odp_unified_reader` | `detect_odp_leaf_reader` | yes / yes |
| detection | `detect_odp_polyglot` | `detect_odp_unified_polyglot` | — (decided by the coordinator's precedence; compare with the predecessor) | yes / — |
| detection | `detect_odp_catalog_alias` | `detect_odp_unified_catalog_alias` | — (decided by the coordinator's precedence; compare with the predecessor) | yes / — |

Predecessors without a new leg:

- `ods_file_{eager,source}_selected_{cell,media}`. These already time the ODS
  leaf owners, and the M0a `File` has no ODS cell or media read
  (`WorkbookView` refuses ODS worksheet access, S5e/S6). Their unified leg
  waits for M1's `get`.
- `docx_source_selected_paragraph`. This is an untimed range-evidence binary
  (S1). The timed DOCX read-one pairs are `docx_semantic_*` and
  `docx_producer_*`.

### Timed scopes that differ from the predecessor

Two predecessors timed a query on a `litchi_docx` document they prepared
outside the timer:

- `docx_semantic_one_paragraph` (`Document::paragraph`)
- `docx_producer_source_selected_paragraph` (`paragraph(i).text()`)

The unified views take no document handle across calls, so both legs of these
pairs also time `Package::document()`. The producer pair first derives the
main story once outside the timer. S10 compares these successors with their
leaf twins, not with the predecessor records.

Every other successor has its predecessor's timed scope. For the
`docx_file_eager_open*` routes, the predecessor recreated the historical eager
facade on the leaf, which also converted core properties at open; the unified
eager owner defers that conversion. `route-map.json` marks each route with
`timed_scope_matches_predecessor` and a note.

## Correctness

Every sample is checked outside the timer, against oracles built from the leaf
crates and the corpus generators, never from the code under measurement.

- **DOCX and PPTX filesystem routes.** The child and the parent check the
  source identity. They then require one projection from each of three opens
  to be identical: the unified eager or byte open, the unified path open, and
  the leaf open. For DOCX the projection covers paragraphs, tables, ordered
  elements, text and metadata. For PPTX it covers slide count and size,
  metadata, every slide's text and name, and the selected slide. The generator
  pins the counts and the timed full text. Source-backed routes on both legs
  carry the predecessor's untimed leaf replay; a classification failure fails
  the sample.
- **XLSX routes.** These must reproduce the typed-oracle semantic digest
  `020fdd14…80627e` and the selected-cell evidence digest `36e53d90…f42ac1`,
  the predecessors' values.
- **In-process routes.** Each sample reproduces the oracle projection, and the
  result reports its digest (`source.unified_ops.projection_sha256`).
- **DOC and PPT routes.** Each leg freezes its outcome before sampling,
  including a refusal. When both legs succeed, their outcomes must be
  identical.

Results carry `source.unified_ops`: the leg, the predecessor, the twin, the
calls inside the timer and a no-claim statement.

## Smoke run

The release build was done once. It uses the crate's default release
profile, not the S1/S10 measurement recipe:

```sh
env CARGO_TARGET_DIR=$PWD/target CARGO_INCREMENTAL=0 \
  cargo build --release --locked --offline \
  --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
```

| Binary property | Value |
|---|---|
| Path | `target/release/litchi-perf-baseline` |
| Size | 71,630,576 bytes |
| SHA-256 | `adc55dd795250c55915e81251771c7d22500be8de7a5ff960774296c980a6646` |
| Toolchain | rustc 1.95.0 (59807616e 2026-04-14) |
| Sources | the committed S8b tree; `crates/` is unchanged from `uo/m0a` `20692d6444` |

Every route ran once from the repository root (`--samples 1 --warmup 0`), in
seven children:

- the DOCX, PPTX and XLSX file groups with `--filesystem-cache warm`;
- the 52 in-process selectors together;
- the DOC and PPT selectors over S1's five fixtures, split as S1 split them:
  - `documentProperties.doc` with `ppt_with_png.ppt`;
  - `FloatingPictures.doc` with `SampleShow.ppt`;
  - `duplicate-style-names.doc` alone, DOC selectors only.

```sh
B=target/release/litchi-perf-baseline
$B --case <the 24 docx_file_{unified,leaf}_* selectors> --warmup 0 --samples 1 \
  --filesystem-cache warm --json OUT/fs_docx.json
# fs_pptx, fs_xlsx and in_process have the same shape
$B --case <the 12 doc_file_* and ppt_file_* selectors> --warmup 0 --samples 1 \
  --ole2-file test-data/ole/doc/documentProperties.doc \
  --ole2-file test-data/ole/ppt/ppt_with_png.ppt --json OUT/ole2_small.json
# ole2_large: FloatingPictures.doc and SampleShow.ppt;
# ole2_refusal: duplicate-style-names.doc with the six doc_file_* selectors
```

The exact argv of every child is in
[`results/0003/s8b/smoke-summary.json`](results/0003/s8b/smoke-summary.json),
which also lists every route with ran yes or no, its result rows and its
corpora. All seven children exited 0, and every route ran and passed its
untimed checks:

| Family | Routes | Ran (yes) | Result rows | Corpora |
|---|---|---|---|---|
| DOCX file | 24 | 24 | 24 | `docx-source-backed-media` |
| PPTX file | 24 | 24 | 24 | `pptx-source-backed-media` |
| XLSX file | 6 | 6 | 6 | `xlsx-cell-values-medium` |
| XLSX bytes | 4 | 4 | 4 | `xlsx-cell-values-medium` |
| DOCX semantic read one | 4 | 4 | 12 | `docx-semantic-large`, `docx-semantic-medium`, `docx-semantic-tiny` |
| DOCX producer read one | 2 | 2 | 2 | `docx-producer-medium` |
| DOC | 6 | 6 | 18 | `doc-facade-real-file` |
| PPT | 6 | 6 | 12 | `ppt-facade-real-file` |
| RTF | 4 | 4 | 12 | `rtf-semantic-plain-large`, `rtf-semantic-plain-medium`, `rtf-semantic-plain-tiny` |
| ODT | 8 | 8 | 8 | `odt-source-backed-repeated-text-media` |
| ODP | 8 | 8 | 8 | `odp-media-textbox-publication` |
| ODS | 4 | 4 | 4 | `ods-media-publication` |
| Detection | 18 | 18 | 54 | `{odt,ods,odp}-{bytes,reader,polyglot,catalog-alias}-{tiny,medium,large}` |
| **Total** | **118** | **118** | **188** | |

DOC and PPT outcomes:

- On `documentProperties.doc`, `FloatingPictures.doc`, `ppt_with_png.ppt`
  and `SampleShow.ppt`, every leg succeeded and the unified and leaf
  outcomes are identical (`legs_agree: true`). For example,
  `FloatingPictures.doc` gives 210 paragraphs, target paragraph 92, and the
  same text digests on both legs.
- Both legs refuse `duplicate-style-names.doc`, as the predecessor facade
  did. `legs_agree` compares the outcome strings, so it is `false` here: the
  two refusals word the error differently. The unified error displays its
  code-level message, `corrupt_file: the file is malformed [Doc]`. The leaf
  names the defect, `invalid stylesheet: style names and aliases must be
  unique`, which the unified error keeps as `Error::leaf()`.

The filesystem evidence matches the predecessors':

- DOCX source replays classify as `catalog-only` for opens and
  `semantic-query` for queries and lifecycles.
- PPTX source replays classify as `catalog-only`, `list-slides` or
  `selected-slide-only`.
- Both DOCX full-text lifecycles return the 10,000-byte oracle text.
- Both XLSX selected-cell routes report the predecessor's evidence digest.

This is not a measurement:

- one sample, no warm-up;
- a shared host, with a 1-minute load of 8.9 at the start and 5.5 at the end;
- a binary built without the measurement recipe.

The raw reports, which hold those single samples, are not kept.

## Gates

`tools/perf-baseline` is a standalone Cargo workspace with its own
`Cargo.lock`. Every cargo command ran with `CARGO_TARGET_DIR=<worktree>/target
CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`,
from the worktree root unless noted.

| Gate | Command | Result |
|---|---|---|
| Format | `cargo fmt --check` (in `tools/perf-baseline`) | clean |
| Clippy | `cargo clippy --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --all-targets -- -D warnings` | clean. It failed at the base commit, on the test helper fixed above. |
| Tests | `cargo test --manifest-path tools/perf-baseline/Cargo.toml --locked --offline` | 643 passed, 1 ignored, 0 failed: lib 593 (12 new), bins 45, `tests/unified_ops_filesystem` 1 (new), `tests/xlsx_filesystem` 4 |
| Allocator target | `cargo check --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --features allocator-metrics --bin litchi-perf-baseline-alloc` | builds |
| Source policy | `python3 -m unittest tools.test_perf_baseline_source_policy` | 28 tests OK (2 new) |
| Coverage index | `python3 -m unittest tools.test_crud_coverage_index` | 50 tests OK; 11 failed before the F-11 fix |
| CI Python list | `python3 -m unittest tools.test_perf_compare tools.test_perf_smoke_baseline tools.test_perf_abba_summary tools.test_perf_abba_package tools.test_perf_resource_profile tools.test_perf_corpus_binding tools.test_validate_opc_serial_eager_allocator_abba tools.test_crud_coverage_index tools.test_validate_perf_default_matrix tools.test_perf_workflow_policy tools.test_perf_baseline_source_policy` | 517 tests OK. The sparse worktree excludes `docs/performance/results/`, so the files these tests read (about 2 MiB) were added to the sparse checkout for the run. |
| Index validator | `python3 tools/validate_crud_coverage_index.py --index docs/performance/crud-coverage-index-v2.json --catalog docs/performance/results/perf-corpus-manifest-v2.json --selector-source tools/perf-baseline/src/lib.rs --checklist docs/CRUD_Scenario_Checklist.md --repo-root .` | 15 categories, 33 mapped selectors |
| Claim registry | `python3 tools/check_perf_claims.py --registry docs/performance/claim-registry-v1.json --repo-root . --mode structural` | 10 claims OK. Strict mode and `tools.test_perf_claims` need evidence trees this sparse worktree lacks; S8b changes neither the registry nor the checker. |
| AGENTS.md pre-commit checks | the two commands in `AGENTS.md` | both print nothing |

## Notes for S10

1. **Two binaries, two selector sets.** The A legs are the pre-M0a binary
   (S1: `litchi-perf-baseline-da4c207478`, SHA-256 `9c887ee5…2c79c04`) running
   the predecessors. The B legs are the post-S9 binary running the unified
   successors. `capture.py` runs one group map per binary, so S10 needs a B
   group map with the successor names. Build it from
   `results/0003/s8b/route-map.json` (`s1_capture_group` keeps S1's grouping,
   including the three OLE2 fixture splits). Keep S1's samples, warm-up, CPU
   pin and `--filesystem-cache warm`, and keep the slow groups (DOCX and PPTX
   file routes) at 30 samples.
2. **Match by route, not by name.** Compare a predecessor with its unified
   successor only where `timed_scope_matches_predecessor` is true. The two
   DOCX read-one routes above are budget-only rows. Use S1's trigger:
   `max(5%, spread_all)`, a 0.5 µs floor below 5 µs, and full separation
   above 10% spread.
3. **The goal's budgets use the pairs in one binary.** Alternate the unified
   and leaf legs ABBA in the post-S9 binary. Open overhead must stay within 5%
   or 20 µs, whichever is larger; the read and query rows are review triggers
   (goal PERFORMANCE BAR).
   - The eager DOCX leaf includes the validation text pass, because the
     unified eager owner runs it.
   - The DOC, PPT and XLSX leaves include the property reads the unified open
     makes. A delta there is the unified layer's work only.
4. **Memory (R5).** DOC and PPT byte routes now retain their input. The
   default harness has no per-route RSS, so S10 needs the
   `litchi-perf-baseline-alloc` build at both revisions if it wants per-route
   memory.
5. **Corpus manifests.** The OLE2 pairs reuse the facade family's corpus
   identity (`doc-facade-real-file`, `ppt-facade-real-file`, same generator).
   A DOC fixture both open paths refuse keeps `entry_count` 0, as before.
   Detection keeps S1's caveat: its byte and reader corpora share an archive,
   so run it without `--corpus-manifest`.
6. **After M1.** The list and read-one unified legs return the
   `litchi::compat` wrappers, which M1 deletes. An M1 change to what a route
   measures needs new names; these names stay frozen with their records.

## Not claimed

- No timing, allocation, I/O or overhead figure. The smoke run proves only
  that each route runs and verifies.
- That the leaf legs are what a typical leaf user writes. They make the leaf
  calls the unified route makes, so that the pair isolates the unified layer.
- That the unified detection leg is the facade's public detector. It is the
  `litchi_ops` coordinator the detector wraps.
