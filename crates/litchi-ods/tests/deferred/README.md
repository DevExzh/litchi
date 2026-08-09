# Deferred ODS integration coverage

The pre-split Rust backlog was adjudicated against the public `litchi-ods`
facade. No Rust source remains deferred. Tests were migrated only where the
current public owner can exercise the original behavior; otherwise the exact
superseding test or the missing post-split capability is recorded below.

| Legacy source | Disposition | Evidence or explicit exclusion |
| --- | --- | --- |
| `legacy/examples/ods_features_demo.rs` | Removed as superseded/inapplicable | Metadata and typed package round-trips are active in `metadata_settings_facade.rs` and `worksheet_facade.rs`. The old console demo, CSV export, and monolithic sheet-selection API are not testable requirements of the split facade. |
| `legacy/examples/ods_reader_test.rs` | Removed as superseded/inapplicable | `worksheet_facade.rs::worksheet_graph_round_trips_repeats_formula_and_style` exercises typed sheet/cell reading; `metadata_settings_facade.rs::facade_round_trips_typed_metadata_and_calculation_settings` exercises metadata. The old arbitrary `test.ods` console smoke test and `sheet_count`/`sheet_by_index`/CSV/text helpers are outside the current API. |
| `legacy/examples/ods_writer_test.rs` | Removed as superseded | `worksheet_facade.rs` actively covers builder sheet/cell/formula authoring and package reopening; `metadata_settings_facade.rs::facade_edits_are_atomic_and_builder_writes_new_parts` covers builder metadata. The remaining file-output narration had no additional assertion. |
| `legacy/examples/write_ods.rs` | Removed as superseded | `worksheet_facade.rs::worksheet_graph_round_trips_repeats_formula_and_style` and `facade_round_trip.rs::builder_and_package_facade_round_trip` cover the asserted builder/package reopen path without an obsolete `tempfile` dependency or CSV helper. |
| `legacy/ods_cell_protection_style_authoring.rs` | Inapplicable after split; source removed | `model::style_protection` exposes inert types, but `Builder`, `Spreadsheet`, and `MutableSpreadsheet` expose no table-cell-style protection catalog or CRUD owner. It cannot be an honest public integration test until that owner exists. |
| `legacy/ods_conditional_formats.rs` | Inapplicable after split; source removed | Conditional-format value types remain under `model::conditional_format`, and `FlatSpreadsheet` parses flat ODS, but neither surface exposes conditional-format attachment, reading, or editing. Parser/authoring integration cannot be reached publicly. |
| `legacy/ods_conditional_style_authoring.rs` | Inapplicable after split; source removed | Conditional style value types exist, but no facade or builder method creates, replaces, removes, or reads cell-style maps. |
| `legacy/ods_conditional_styles.rs` | Inapplicable after split; source removed | The real LibreOffice fixture remains available, but the current spreadsheet facade has no conditional-cell-style lookup. A fixture test would have no public observation point. |
| `legacy/ods_drawing_style_resources.rs` | Inapplicable after split; source removed | The current `drawing` facade exports `Frame` and `Part`; it has no fill-image, gradient, hatch, marker, opacity, or stroke-dash catalog methods. |
| `legacy/ods_grid_padding.rs` | Partially migrated | Packaged-fixture bounded parsing and mutable preservation are active in `tests/grid_padding.rs`. Flat `.fods` parsing and editing are covered in `tests/deferred_flat_spreadsheet.rs`; the legacy grid-padding-specific flat cases have no focused replacement. |
| `legacy/ods_hyperlinks.rs` | Inapplicable after split; source removed | The current worksheet `Cell` has no hyperlink collection. Inert rich-text/link byte preservation remains active in `source_features.rs::no_op_package_round_trip_preserves_compact_rich_text_xml_exactly`, but semantic hyperlink enumeration is not publicly exposed. |
| `legacy/ods_sheet_shapes.rs` | Inapplicable after split; source removed | The public drawing surface has frames only; `Shape`, `SheetShape`, typed anchors, shape CRUD, and flat-document rewriting are absent. Image-frame discovery is separately active in `facade_round_trip.rs`. |
| `legacy/ods_sparklines.rs` | Inapplicable after split; source removed | Sparkline value types remain public under `model::sparkline`, but worksheets and `FlatSpreadsheet` expose no sparkline-group parse, attachment, or CRUD API. |
| `legacy/ods_tracked_change_authoring.rs` | Removed as exactly superseded | `tracked_changes.rs::reads_all_four_families_and_authors_odf_14_order_with_schema_prefixes` covers all four variants and authoring; `transaction_crud_reorder_acceptance_reopen_and_rollback_are_atomic`, `duplicate_unknown_forward_wrong_family_and_cyclic_references_are_rejected`, and `invalid_cross_record_operations_fail_immediately_without_draft_mutation` cover atomic mutation and invalid references. |
| `legacy/sheet_images.rs` | Removed as superseded/inapplicable | `facade_round_trip.rs::spreadsheet_facade_discovers_resources_and_extracts_local_images` covers package and linked image discovery plus inert byte access. Sheet image insertion, numbering, and CRUD are explicitly excluded because the split facade exposes discovery only. |

External links, DDE sources, and linked images remain inert: active tests inspect
their metadata or preserve their XML without fetching external resources.

The five cross-family/flat deferred sources are now covered for their ODS cases
by `deferred_ods_corpus.rs`, `deferred_flat_spreadsheet.rs`, and
`deferred_ods_hardening.rs`. The ODP/ODG/ODC/ODM-only hardening source remains
owned by those family crates. Flat cell edits preserve unrelated XML and refuse
unmodeled changed-row markup. The legacy behavior that invented an
`office:meta` section during an unrelated cell edit is intentionally not
retained: preserve-mode editing neither repairs nor normalizes untouched parts.
