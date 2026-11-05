# S8c — facade tests on `litchi::open` and `File`

- Status: done; tests only, no library code changed.
- Date: 2026-10-05
- Branch: `uo/m0a-s8c-tests`, based on `uo/m0a` at `20692d6444` (S6, S6b and
  S7 merged).
- Scope: M0a step S8c (record 0003 §4.4 and §5.2). The facade's integration
  tests (`crates/litchi/tests/`) and its fuzz crate move from the
  format-family types (`litchi::Document`, `litchi::Presentation`,
  `litchi::sheet::Workbook`, the `open_workbook*` and `open_xlsb_workbook*`
  openers) to `litchi::open*` and `File`, keeping every assertion, so that S9
  can delete those types.

After this step, `tests/markdown_file_parity.rs` is the only facade test that
names the format-family types. It compares them with `File` on purpose, and
it now also pins the golden files that carry its semantics past S9 (§4).

## 1. How the tests map

Unless a row says otherwise, a ported test keeps its name, fixtures and
expected values. The general rules:

| Old form | New form |
|---|---|
| `Document::open(p)` / `Presentation::open(p)` / `Workbook::open(p)` | `litchi::open(p)` |
| `*::from_bytes(b)` | `litchi::open_bytes(b)` (source-backed over the bytes, as the facade's byte route already was) |
| `*::open_with_limits(p, l)` / `*::from_bytes_with_limits(b, l)` | `litchi::open_with(p, &OpenOptions::new().ooxml_limits(l))` / `open_bytes_with` |
| `UnexpectedFormat { detected }` from a family opener | the family view (`file.document()`, `.presentation()`, `.workbook()`) returns `Code::UnsupportedFormat` with `error.detected() == Some(format)` |
| `NotOfficeFile` | `litchi::open` returns `Code::UnsupportedFormat` with `detected() == None` |
| `ResourceLimit { resource, observed, limit }` | `Code::LimitExceeded`; resource, observed value and limit are asserted on the typed source `error.leaf()` (`Source::Opc(OpcError::ReadLimit { resource, actual, maximum })`) |
| `SourceChanged` (core error, or boxed in a sheet error) | `Code::IoError`; the leaf is a core or OPC `SourceChanged` |
| `workbook.sheet(selector)` | `File::Xlsx(xlsx) => xlsx.sheet(selector)` (`litchi::format::xlsx`) |
| `Box<dyn WorkbookTrait>` from `open_*_workbook*_dyn` | `file.workbook()?`, the `WorkbookView` with its static `WorkbookTrait` |
| an error boxed by a sheet trait method (`litchi_core::Error`, `xlsb::package::Error`) | a boxed `litchi::Error` whose leaf is that typed error |
| `document.to_markdown()` | `file.to_markdown()` (`impl ToMarkdown for File`) |
| `litchi::document::Table` | `litchi::compat::word::Table` |
| `litchi::detection_smart::{DetectedFormat, detect_format_smart*}` | `litchi::common::detection::{...}`: the same items, through the facade's documented detection path |

### `unexpected_format.rs` (16 tests)

| Test | New form; semantics note |
|---|---|
| `an_xlsx_given_to_the_document_opener_reports_xlsx` | `open(styles.xlsx)?.document()` refused, `detected == Xlsx` |
| `a_pptx_given_to_the_document_opener_reports_pptx` | `.document()` on the PPTX, `detected == Pptx` |
| `a_legacy_ole2_presentation_given_to_the_document_opener_reports_ppt` | `.document()` on the PPT, `detected == Ppt` |
| `an_odf_presentation_given_to_the_document_opener_reports_odp` | `.document()` on the ODP, `detected == Odp` |
| `a_non_office_file_given_to_the_document_opener_still_reports_not_an_office_file` | `open(lena.png)` refused, `detected == None` |
| `a_docx_given_to_the_presentation_opener_reports_docx` | `.presentation()` on the DOCX, `detected == Docx` |
| `a_legacy_ole2_document_given_to_the_presentation_opener_reports_doc` | `.presentation()` on the DOC, `detected == Doc` |
| `an_odf_document_given_to_the_presentation_opener_reports_odt` | `.presentation()` on the ODT, `detected == Odt` |
| `a_non_office_file_given_to_the_presentation_opener_still_reports_not_an_office_file` | as the document case |
| `a_docx_given_to_the_workbook_opener_reports_docx` | `.workbook()` on the DOCX, `detected == Docx` |
| `a_pptx_given_to_the_workbook_opener_reports_pptx` | `.workbook()` on the PPTX |
| `a_legacy_ole2_document_given_to_the_workbook_opener_reports_doc` | `.workbook()` on the DOC |
| `an_odf_document_given_to_the_workbook_opener_reports_odt` | `.workbook()` on the ODT |
| `an_odf_presentation_given_to_the_workbook_opener_reports_odp` | `.workbook()` on the ODP |
| `a_non_office_file_given_to_the_workbook_opener_still_reports_not_an_office_file` | as the document case |
| `the_workbook_opener_separates_every_wrong_format_it_can_name` | the five refusals report five distinct `detected` formats. Stronger: the old test deduplicated the expected list, the new one the reported values. |

### `read_limits_facade.rs` (8 tests)

| Test | New form; semantics note |
|---|---|
| `docx_limits_are_nameable` | `litchi::docx::ReadLimits` is accepted by `OpenOptions::ooxml_limits` (the same OPC type), `open_with` / `open_bytes_with` / `common::detection::detect_format_smart_with_limits` are nameable |
| `document_limits_bound_path_and_byte_ingress` | default open succeeds (format DOCX); a limit one byte short refuses the path and byte routes. Stronger: asserts `LimitExceeded` and `ReadLimit { InputBytes, actual = len, maximum = len - 1 }`, where the old test asserted `is_err()` only. |
| `document_limits_reject_invalid_suffix_and_unknown_extension_ooxml` | the 32-byte `.docx` under an 8-byte limit and the extension-less DOCX under `len - 1` are refused with the exact input limit; `detect_ooxml_format_with_limits` still declines |
| `pptx_limits_are_nameable` | as DOCX |
| `presentation_limits_bound_path_and_byte_ingress` | as DOCX |
| `presentation_limits_reject_invalid_suffix_and_unknown_extension_ooxml` | as DOCX |
| `xlsx_limits_and_bounded_sheet_helpers_are_nameable` | the bounded sheet openers are the unified openers with `ooxml_limits`. New: `SimpleNormal.xlsx` one byte over the limit is refused on both routes. |
| `xlsb_limits_and_bounded_sheet_helpers_are_nameable` | as XLSX, over `Simple.xlsb`, on the path and byte routes in both `SourceMode::Auto` and `SourceMode::Eager` (the facade had eager and trait-object openers) |

### `unified_sheet_selectors.rs` (5 tests)

| Test | New form; semantics note |
|---|---|
| `public_selector_preserves_exact_owned_cell_semantics` | `open_bytes` + `xlsx.sheet("dAtA")`: case-folded selection, positions, cell views and ranges unchanged; coordinate and range errors are `Source::Xlsx(Coordinate / Range)` leaves |
| `selected_access_defers_unselected_malformed_worksheet` | unchanged semantics through `xlsx.sheet` |
| `non_grid_sheet_returns_typed_not_worksheet_error` | leaf `XlsxError::NotWorksheet { sheet: "Chart" }`; also pins `Code::UnsupportedType`, which `SelectedWorksheet::cell` documents |
| `non_xlsx_workbook_returns_explicit_unsupported_error` | **form changed** (§3, item 1): an XLSB opens as `File::Xlsb`, whose handle has no `sheet` method; the test asserts that and that the XLSB workbook view reads |
| `selected_handle_preserves_source_change_error` | `open(path)`; after the rewrite, both the selected cell read and a new selection report `IoError` with a source-change leaf |

### `catalog_detection_arbitration.rs` (9 tests)

Detection assertions are unchanged; only the opens moved.

| Test | New form; semantics note |
|---|---|
| `document_limit_arbitration::ordinary_odt_uses_native_policy_but_ooxml_suffix_honors_docx_input_limit` | ordinary ODT bytes and `.odt` path under a 1-byte OPC limit open as `File::Odt`; the `.DOCX`-suffixed ODT is refused with `ReadLimit { InputBytes, len, 1 }`; without limits it opens as ODT with text `bounded`; both ODT/DOCX polyglots (`[Content_Types].xml` and lower-case) are refused on the bytes and path routes with the exact input limit |
| `content_type_matching::accepts_canonical_content_type_casing_but_rejects_prefix_and_suffix_near_matches` | unchanged (detection only) |
| `ordinary_ods_detection::ordinary_ods_remains_detectable_without_an_opc_catalog` | unchanged; `DetectedFormat`/`detect_format_smart` imported from `common::detection` |
| `ordinary_odp_detection::ordinary_odp_remains_detectable_without_an_opc_catalog` | as ODS |
| `ooxml_odf_arbitration::ooxml_precedence_survives_an_odf_mimetype_marker` | `open_bytes` of the polyglot is `File::Pptx` with one slide (`presentation()?.slide_count()`) |
| `ooxml_odf_arbitration::ordinary_odp_bytes_use_native_policy_before_pptx_limits` | `open_bytes_with` under a 1-byte OPC limit is `File::Odp` with one slide |
| `ooxml_odf_arbitration::tight_limits_do_not_fall_back_to_the_odp_marker` | detection unchanged; the open is refused with `LimitExceeded`, leaf `ReadLimit { ContentTypesBytes, maximum: 1 }` (§3, item 3) |
| `ooxml_odf_arbitration::renamed_odp_with_tiny_pptx_limit_uses_the_odp_policy` | `open` and `open_with` (1-byte OPC limit) of the `.pptx`-named ODP are `File::Odp` with one slide |
| `wrong_family_workbook::wrong_family_ooxml_path_reports_the_detected_family` | **fixture added** (§3, item 2): a readable one-slide PPTX at an `.xlsx` path is refused by `.workbook()` with `detected == Pptx`; the original catalog-only package is refused at open with `CorruptFile` and `format == Pptx` (in a build without `pptx`: `UnsupportedFormat`, `detected == Pptx`) |

### `catalog_first_detection.rs` (4 tests)

All four (`catalog_first_detection_matches_bytes_reader_and_typed_docx_without_payload_reads`, `typed_catalog_and_owner_restore_reader_position_on_error_and_limit`, `typed_malformed_package_error_is_distinct_from_fresh_no_match`, `smart_detection_still_returns_a_usable_eager_docx_variant`) are detection tests and were unchanged apart from importing `DetectedFormat` and `detect_format_smart` from `litchi::common::detection` instead of `litchi::detection_smart`.

### `tolerance_report_facade.rs` (12 tests)

Every test opens its fixtures with `litchi::open` (and `litchi::open_bytes` where the old test also used `from_bytes`) and calls `file.tolerance_report()`; fixtures, defect classes, counts and clean reports are unchanged:
`doc_repairs_are_reported_through_the_document_facade`,
`docx_core_properties_repairs_are_reported_through_the_document_facade` (path and bytes; one `MissingDateType`, twice),
`rtf_repairs_are_reported_through_the_document_facade`,
`odt_repairs_are_reported_through_the_document_facade`,
`pages_reports_are_untracked` (unchanged: Pages has no `File` variant, so it keeps `ToleranceReport::Untracked(Pages)`),
`ppt_repairs_are_reported_through_the_presentation_facade`,
`pptx_core_properties_repairs_are_reported_through_the_presentation_facade`,
`odp_repairs_are_reported_through_the_presentation_facade`,
`xls_repairs_are_reported_through_the_workbook_facade`,
`xlsx_core_properties_repairs_are_reported_through_the_workbook_facade` (the metadata check reads `file.metadata()?` as `Option`, which is `Some` for XLSX),
`xlsb_core_properties_repairs_are_reported_through_the_workbook_facade`,
`ods_repairs_are_reported_through_the_workbook_facade`.

### Markdown: `markdown_docx_lists.rs` (2), `markdown_docx_semantics.rs` (10), `markdown_real_fixtures.rs` (1)

| Test | New form; semantics note |
|---|---|
| `renders_resolved_numbering_semantics` | `open_bytes` and `open(path)`, same golden |
| `rejects_missing_numbering_definition` | refused on the byte route, and now also on the path route (from `markdown_file_parity`) |
| `package_semantics_have_compact_golden_markdown` | same golden on both routes; after the source mutation, `IoError` with a source-change leaf |
| `parallel_option_preserves_serial_docx_markdown_output` | unchanged, plus the exact `First\n\nSecond\n\n` (from `markdown_file_parity`) |
| `source_path_table_projection_matches_eager_markdown` | path, bytes and `SourceMode::Eager` all give tables `["A|B|C|D|"]` and the Markdown the facade's `Document` produced (recorded constant `TABLE_MARKDOWN`). The old test compared the path route with the facade byte route only, which was source-backed too, so the eager owner the name refers to was not exercised. |
| `source_path_mce_projection_matches_eager_markdown` | all three routes give text `fallback` and the recorded Markdown `fallback\n\n` |
| `unsupported_link_placement_metadata_is_refused` | on all three routes: `Code::UnsupportedElement`, message contains `w:tgtFrame`, format DOCX (the facade returned `Unsupported(message)`) |
| `mixed_text_and_drawing_run_is_refused` | as above, for `interleaved DOCX text and embedded semantics` |
| `embedded_ole_object_projects_its_preview_image`, `embedded_ole_object_with_unresolvable_preview_degrades_to_alt_text`, `floating_drawing_projects_as_image`, `bookmark_boundaries_are_invisible_to_markdown` | same goldens, now on all three routes (the old tests used the byte route only) |
| `real_docx_plain_paragraphs_have_compact_golden_markdown` | `litchi::open(p)?.to_markdown()`, returning `litchi::Result` |

### `xlsb_facade.rs` (23 tests)

The facade's lazy path and byte routes are `litchi::open` / `litchi::open_bytes`; its trait object is `file.workbook()?`; its eager byte entry `open_xlsb_workbook_from_bytes` is both `SourceMode::Eager` and the leaf call it wrapped (`xlsb::Package::from_slice(..)?.into_workbook()`). Each test opens its input with `File`, and asserts `format() == Xlsb`.

| Test | New form; semantics note |
|---|---|
| `path_facade_defers_malformed_unselected_worksheet` | `open(path)`; names and count read, `text()` fails |
| `owned_bytes_facade_defers_malformed_unselected_worksheet` | `open_bytes` |
| `dynamic_owned_bytes_facade_defers_malformed_unselected_worksheet` | `open_bytes(..)?.workbook()?` through `WorkbookTrait` |
| `dynamic_owned_bytes_facade_retains_input_after_caller_mutation` | `open_bytes(bytes.as_slice())` copies once; zeroing the caller's buffer leaves `Data!A1 == "DATA"` |
| `dynamic_owned_bytes_facade_enforces_exact_input_limit` | `open_bytes_with(ooxml_limits)`: the exact limit is accepted (2 worksheets); one byte less is refused with `ReadLimit { InputBytes, len, len - 1 }` (old: any `litchi_core::Error`) |
| `dynamic_owned_bytes_facade_enforces_exact_part_limit` | as above for `max_part_bytes`: `ReadLimit { PartBytes, maximum: largest - 1 }` |
| `dynamic_xlsb_byte_entrypoint_rejects_xlsx` | **form changed** (§3, item 6): `open_bytes` routes the XLSX to `File::Xlsx`; the XLSB leaf entry refuses it |
| `dynamic_xlsb_byte_entrypoint_rejects_arbitrary_bytes` | `open_bytes` refuses with `UnsupportedFormat`, `detected == None` |
| `path_facade_reports_typed_source_change_for_catalog_and_reads` | after the append: `worksheet_names`, `worksheet_count` and `text` report `IoError` with a source-change leaf |
| `lazy_facade_catalog_matches_eager_xlsb_tabs_and_date_system` | path view names, count and 1904 flag equal the leaf's; the eager facade's 1904 flag now comes from `SourceMode::Eager` (the old "eager" trait object was the source-backed path opener) |
| `source_backed_facade_text_matches_eager_workbook` | `open(path)?.text()` equals the leaf-derived text |
| `source_backed_facade_skips_chart_tabs_when_selecting_worksheets` | through the path view's `WorkbookTrait`, unchanged |
| `source_backed_facade_text_skips_chart_tabs` | `text() == "DATA\nTAIL\n"` |
| `source_backed_facade_strictly_refuses_sparkline_materialization` | path and byte views refuse the tail cell, and `text()` refuses, each with `UnsupportedCapability` and leaf `PackageError::UnsupportedFeature("… sparkline groups …")`; the leaf eager workbook reads it, and so does the new `SourceMode::Eager` file (`text` and the tail cell) |
| `eager_xlsb_byte_entrypoint_rejects_malformed_sheet_while_lazy_facade_opens` | lazy path names read; the leaf and `SourceMode::Eager` both refuse at open |
| `eager_and_dynamic_facade_select_active_tail_in_all_worksheet_catalog` | leaf and byte view agree: active index 1, `Tail` |
| `eager_and_dynamic_facade_map_active_tail_catalog_position_to_worksheet_ordinal` | as above, plus the leaf round trip |
| `eager_and_dynamic_facade_reject_active_chart_as_active_worksheet` | leaf unchanged; the byte view's boxed `litchi::Error` keeps the leaf `UnsupportedFeature("XLSB active sheet is not a worksheet")` |
| `eager_xlsb_rejects_out_of_range_active_catalog_position` | leaf unchanged; the byte and path routes now both refuse at open with the typed leaf `InvalidFormula("… BrtBookView sheet position …")`. Stronger: the old byte route only showed "a core error", because the facade flattened it to a string. |
| `xlsb_writer_default_active_tab_is_deterministic` | leaf-only test, unchanged |
| `path_facade_selects_active_tail_after_chart_tab_and_observes_source_changes` | path view; after the append, `active_worksheet` reports a source change |
| `path_facade_reports_typed_error_for_active_chart_and_source_change_takes_precedence` | path view; the typed active-chart refusal, then the source change |
| `chart_first_writer_defaults_active_catalog_to_first_worksheet` | leaf and byte view agree: active index 0, `Data` |

### `pages_facade.rs` (1 of 7 tests)

`body_storage_reaches_facade_paragraphs` read Pages through the facade's `Document`. It now reads `litchi::iwork::Document::open(path)?.snapshot()` (under `iwork`): exactly one text item, equal to the native text storage, with role `SectionAdditional`; one body section with no plain paragraphs and that storage as its additional text. New, per §4.4: `litchi::open` on the `.pages` file returns `UnsupportedFormat`, `detected == Pages`, with a suggestion naming `litchi::iwork::Document::open` (in every build with `pages`). The other six tests use `litchi::pages` and are unchanged.

### `unified_open_facade.rs` (5 of 14 tests)

S7's parity tests compared `File` with the format-family types directly. They now compare with the goldens recorded from those types (§4):

| Test | New form; semantics note |
|---|---|
| `the_root_error_is_the_unified_error_and_the_core_error_stays_nameable` | `open` of a missing file is `FileNotFound`; `litchi::common::Error` stays nameable and converts into `litchi::Error` (`FileNotFound`, leaf `Source::Core(Io)`). The half that said the format-family types report `common::Error` describes the deleted types (§3, item 4). |
| `document_tolerance_reports_come_from_the_handles` | the 14 recorded rows (7 fixtures × path and bytes, metadata read first) equal `format!("{:?}", file.tolerance_report()?)`; the DOC bookmark assertion is unchanged |
| `presentation_tolerance_reports_come_from_the_handles` | 4 recorded rows; the PPT repair assertion is unchanged |
| `workbook_tolerance_reports_come_from_the_handles` | 7 recorded rows (text and metadata read first) |
| `pptx_slide_catalog_and_size_come_from_the_handle` | all 55 recorded `ooxml/pptx` fixtures: catalog `(position, slide id)` and slide size equal the `Pptx` handle's. Stronger: a signed package is compared through `Signed::file()`; S7 skipped those. |

### `eval_xlsx_integration` (3 tests + 1 new)

`eval_aggregates_and_logical`, `eval_pduration_and_rri` and `eval_lookup_and_text_functions` evaluate through `FormulaEvaluator::new(&litchi::open(p)?.workbook()?)` instead of the boxed `open_workbook(p)`; values unchanged. This is the S6 follow-up: `litchi_eval::FormulaEvaluator<W: WorkbookTrait + Send + Sync + ?Sized>` reads `WorkbookView` directly. New: `routes::eval_reads_the_workbook_view_of_every_route` builds a workbook with `SUM(A1:A2)*A2` and evaluates it to 15 through the path, byte and `SourceMode::Eager` views.

### In-source test outside `tests/`: `src/detection_smart/prepared_index_tests.rs` (7 tests)

This module is outside the files S9 deletes but opened through the format-family types, so it would have blocked S9. Ported:

| Test | New form; semantics note |
|---|---|
| `{docx,pptx,xlsx}_detect_then_open_indexes_twice_and_prepared_open_indexes_once` | `detect_file_format` + `crate::open` makes 2 OPC probes; `detect_and_prepare` + `crate::open_prepared` makes 1; the opened format is asserted |
| `a_prepared_handle_refuses_a_source_rewritten_before_adoption` | `open_prepared` after the rewrite: `IoError`, source-change leaf |
| `a_prepared_docx_handle_is_refused_by_the_presentation_opener` | `open_prepared` adopts the DOCX; `.presentation()` refuses with `detected == Docx` |
| `non_ooxml_inputs_are_still_classified_without_a_prepared_handle`, `a_non_office_file_is_rejected_by_both_detectors` | unchanged (detection only) |

### Unchanged targets

`encryption_facade`, `vba_inspection_facade`, `semantic_facade`, `semantic_reader_facade`, `numbers_facade`, `numbers_semantic_values`, `numbers_semantic_view`, `keynote_facade` and every `iwork_*` target use leaf or `litchi::iwork` APIs only (their `Document` is `litchi::iwork::Document`).

## 2. Moved to leaves

Nothing moved to a leaf crate. Tests that are about a leaf stay in the facade on the facade's leaf re-exports, as before: `xlsb_writer_default_active_tab_is_deterministic` and the leaf halves of the XLSB active-sheet tests (`litchi::xlsb::Workbook`), the encryption and VBA name checks, and the leaf eager oracle `xlsb::Package::from_slice(..).into_workbook()` that replaces `open_xlsb_workbook_from_bytes`.

## 3. Assertions that changed form or could not be expressed

1. **`non_xlsx_workbook_returns_explicit_unsupported_error`.** The facade refused `Workbook::sheet` on an XLSB at run time (`Unsupported("… only for XLSX")`). In the unified API the selector is a method of the `Xlsx` handle, so an XLSB, which opens as `File::Xlsb`, cannot call it: the refusal is a compile-time guarantee. No call site returns `unsupported_capability` for it, so plan §4.4's "on a non-XLSX workbook → `unsupported_capability`" cannot be asserted. The test asserts the `File::Xlsb` variant and a working workbook view instead.
2. **`wrong_family_ooxml_path_reports_the_detected_family`.** The facade's workbook opener refused the catalog-only PPTX from its catalog. `litchi::open` builds the PPTX owner the catalog selects, and that owner refuses the dangling slide reference (`CorruptFile`, `format == Pptx`) before any view exists. Both halves are kept: the refusal through `.workbook()` uses a readable one-slide PPTX, and the original fixture asserts the corrupt-file error that still names PowerPoint.
3. **`tight_limits_do_not_fall_back_to_the_odp_marker`.** The old test asserted `resource == InputBytes, limit == 1` because the facade mapped every byte-sized OPC limit to `Resource::InputBytes` (`presentation/prs.rs` `map_opc_read_resource`). The typed leaf names the limit that refused the package, `ContentTypesBytes`, with `maximum == 1`.
4. **`the_root_error_is_the_unified_error_and_the_core_error_stays_nameable`.** "The format-family types report `common::Error`" is a statement about the deleted types. What remains is kept: `common::Error` is nameable and converts into the unified error.
5. **`body_storage_reaches_facade_paragraphs`.** `litchi::iwork` has no `elements()` / `DocumentElement::as_paragraph`. The snapshot's single text item (and the section's single storage) stands for the facade's single paragraph element.
6. **`dynamic_xlsb_byte_entrypoint_rejects_xlsx`.** There is no XLSB-only opener in the unified API. Detection routes the XLSX to `File::Xlsx` (asserted), and the XLSB leaf entry rejects it (asserted).
7. **Nameability checks** of `open_with_limits`, `from_bytes_with_limits` and the `sheet::open_*_with_limits` openers. Those functions are deleted. Their replacements, `open_with`, `open_bytes_with` and `OpenOptions::ooxml_limits` taking each leaf's `ReadLimits` type, are named instead, and exercised at run time.
8. **Read limits "through `Error::source()`"** (plan §4.4): asserted through `Error::leaf()`, the typed `Source` enum. `std::error::Error::source()` returns the same leaf error, untyped.

## 4. Golden files

All three are under `crates/litchi/tests/golden/`, recorded from the format-family types by `markdown_file_parity.rs`:

| File | Bytes | Rows | Content | Read by |
|---|---|---|---|---|
| `markdown_corpus.tsv` | 39,257 | 438 fixtures × 6 columns | per DOCX/DOC/RTF/ODT/PPTX/PPT/ODP fixture under `test-data`: for `default`, `plain` (`with_styles(false)`) and `metadata+html` (`with_metadata(true)` + `StyledHtml`), on the path and byte routes, `<length>:<FNV-1a 64>` of the facade Markdown or `!<facade error variant>` | `markdown_corpus_golden.rs` (new): `File` matches the digest where the facade produced Markdown, and refuses where it refused (2,076 identical, 552 refused by both, as S7's parity) |
| `tolerance_reports.tsv` | 4,650 | 25 | family, route, fixture, `{:?}` of the facade report | `unified_open_facade.rs` tolerance tests |
| `pptx_slide_catalog.tsv` | 2,858 | 55 | file name, slide width and height, `position:slide id` list | `unified_open_facade.rs` catalog test |

SHA-256 at this commit:

```text
8c04603c257d7bdd29b77788bd9e66ee309bd0a6f8597daec550ef4ada90039d  markdown_corpus.tsv
64f5a7ce551f25c4e8dbebe8c58ff89cb0586e33c7337bf63dafff14f17e2119  tolerance_reports.tsv
99ace97b39b26d3492024dcb8a16736e661fe4d2bc566285c8e2b1be30917fcf  pptx_slide_catalog.tsv
```

Two goldens are inline constants instead: the table and MCE documents' Markdown in `markdown_docx_semantics.rs`, captured from the facade's `Document` on 2026-10-05.

How they were recorded and how they stay honest until S9: `markdown_file_parity.rs` gained three tests, `the_markdown_corpus_golden_is_the_facade_types_output`, `the_tolerance_report_golden_is_the_facade_types_output` and `the_pptx_slide_catalog_golden_is_the_facade_types_output`. Each recomputes its golden from the format-family types and requires the committed file to equal it. With `LITCHI_BLESS_GOLDEN=1` they rewrite the files instead:

```sh
LITCHI_BLESS_GOLDEN=1 cargo test -p litchi --all-features --test markdown_file_parity golden
```

FNV-1a is used because the facade has no hashing dependency. The digest plus length detects a regression; to see the Markdown itself, run the facade types on a pre-S9 commit.

## 5. Notes for S9

1. **Delete `tests/markdown_file_parity.rs` whole.** Every `File`-side assertion in it has a home that survives:
   - its fixtures and goldens: `markdown_docx_lists.rs`, `markdown_docx_semantics.rs` and `markdown_real_fixtures.rs`, which now cover both routes wherever the parity test did;
   - its corpus test: `markdown_corpus_golden.rs`.

   Its old-versus-`File` and golden-versus-old checks go with the old types. After that, the goldens can only be re-blessed from `File`, by hand, so treat a golden diff as a behaviour change to review, not to bless.
2. **Facade paths the ported tests rely on.** Keep them, or re-point the tests:
   - `litchi::common::detection::{DetectedFormat, detect_format_smart, detect_format_smart_with_limits, detect_file_format_from_bytes, detect_format_from_reader}`;
   - `litchi::detection_smart::ooxml::{detect_ooxml_format_with_limits, detect_ooxml_format_from_bytes_with_limits}` (`read_limits_facade.rs`); `litchi_ops::detect::ooxml` is the same module;
   - `litchi::sheet::{CellValue, FormulaEvaluator, WorkbookTrait}` (eval and XLSB tests), which survive S9's `sheet` deletions only if `sheet/mod.rs` keeps those re-exports;
   - `litchi::markdown::{MarkdownOptions, TableStyle, ToMarkdown}`, `litchi::common::Error`, `litchi::ToleranceReport`;
   - `crate::detection_smart::{detect_file_format, detect_and_prepare, opc_probe_count, reset_opc_probe_count}` (`prepared_index_tests.rs`). The probe counter is `litchi_ops::__migration::detect::*` today, so it needs a public test-support path when `__migration` goes.
3. **In-source tests.** `src/detection_smart/prepared_index_tests.rs` is ported (§1). The in-source tests of the facade's old Markdown adapter (`src/markdown/document.rs`, `src/markdown/writer.rs`) all have same-named counterparts in `litchi-ops/src/markdown/`; delete them with the adapter. The tests under `src/document/`, `src/presentation/` and `src/sheet/{workbook,adapters,…}` go with their files (S6 ported them to `File` tests).
4. **Fuzz.** No target in `crates/litchi/fuzz` uses the format-family types; nothing to port. `cargo check --manifest-path crates/litchi/fuzz/Cargo.toml --bins --keep-going` passes 57 of 58 bins. The exception is `keynote_chart_title`, which fails with E0283 (`edit.set(distinct_title(..).as_ref())` cannot infer the `AsRef` target). That failure predates this step: the fuzz crate builds the non-test `litchi`, which this step did not change, and the target file is unchanged since `5533b7e3b3` (2026-08-23). It is iWork fuzz code, which this program does not edit, so it is left for the iWork owners. The fuzz crate's generated `Cargo.lock` is git-ignored.
5. Test names keep the word "facade" (for example `…_through_the_document_facade`). It now means the `litchi` crate's entry points, not the deleted types; the names were kept so this record maps one to one.

## 6. Verification

On this branch (`CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`):

- `cargo fmt --check -p litchi`: clean.
- `cargo clippy -p litchi --all-targets --all-features -- -D warnings`: clean (re-run after a `touch` of `src/lib.rs`, so every target was linted).
- `RUSTDOCFLAGS=-Dwarnings cargo doc -p litchi --no-deps --all-features`: clean.
- Each test target run alone with `cargo test -p litchi --all-features --test <target>`:

| Target | Passed | | Target | Passed |
|---|---|---|---|---|
| catalog_detection_arbitration | 9 | | markdown_real_fixtures | 1 |
| catalog_first_detection | 4 | | numbers_facade | 5 |
| encryption_facade | 2 | | numbers_semantic_values | 2 |
| eval_xlsx_integration | 4 | | numbers_semantic_view | 1 |
| iwork_error_model | 2 | | pages_facade | 7 |
| iwork_facade | 8 | | read_limits_facade | 8 |
| iwork_merge_readers | 3 | | semantic_facade | 3 |
| iwork_merge_transactions | 3 | | semantic_reader_facade | 3 |
| iwork_numbers_compatibility | 3 | | tolerance_report_facade | 12 |
| iwork_pages_table_deletion | 3 | | unexpected_format | 16 |
| iwork_path | 7 | | unified_open_facade | 14 |
| iwork_source_errors | 2 | | unified_sheet_selectors | 5 |
| keynote_facade | 12 | | vba_inspection_facade | 5 |
| markdown_corpus_golden | 1 | | xlsb_facade | 23 |
| markdown_docx_lists | 2 | | `--lib` | 352 |
| markdown_docx_semantics | 10 | | `--doc` | 49 (2 ignored) |
| markdown_file_parity | 12 | | | |

  That is 31 targets with 192 tests, all passing, and none failing or ignored.
- No ported test references the old types. Paths and the `open_*workbook*`
  openers:

  ```sh
  grep -rnE 'litchi::(Document|Presentation|Workbook|DocumentElement|SlideDescriptor)\b|litchi::sheet::(Workbook\b|Selected|functions)|litchi::(document|presentation)::|open_(xls_|xlsb_)?workbook' crates/litchi/tests \
    | grep -v '^crates/litchi/tests/markdown_file_parity.rs'
  ```

  prints nothing (exit 1). Brace imports, across lines:

  ```sh
  grep -rlPz 'use litchi::(sheet::)?\{[^}]*\b(Document|Presentation|Workbook|DocumentElement|SlideDescriptor)\b' crates/litchi/tests
  ```

  prints only `crates/litchi/tests/markdown_file_parity.rs`. Run as
  `git grep` against the base `20692d6444`, the same two patterns flag 14
  files (catalog_detection_arbitration, the three eval modules,
  markdown_docx_lists, markdown_docx_semantics, markdown_real_fixtures,
  pages_facade, read_limits_facade, tolerance_report_facade,
  unexpected_format, unified_open_facade, unified_sheet_selectors and
  xlsb_facade, plus the parity file).
