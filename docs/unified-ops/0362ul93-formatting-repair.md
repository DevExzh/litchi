# 0362ul93 — scoped workspace formatting repair

- Date: 2026-10-25.
- Base: integration `88ca2fd38fdd23ca2ee4ad22449a83893804e4e1`.
- Status: source-ready; the coordinator must rerun complete workspace formatting after all pending source merges.

## Captured scope

The preserved actual formatter failure is `(deleted-coord-artifact)/00-fmt-check-attempt-actual-f5-00.log` (949,997 bytes; SHA-256 `6444246d15f665bdb857146323f77dc64fb61e7542e483b7c1f5cfa4ffe23b53`). It was the only executed row and failed before any subsequent validation row. I extracted each Diff in path, removed the captured checkout prefix, sorted and deduplicated it. The resulting 335 tracked Rust paths are bound by this newline-delimited path-list SHA-256:

`6b53128af05235f8aef4bb89b0898fb743215b470af59acbd880fe0e09d2ebb5` (112 `vendor/` paths, 223 workspace paths; 1,907 repeated diff-hunk lines in the capture). This digest is SHA-256 over the sorted relative path strings, each followed by one LF byte.

The selected path list was checked against the fresh leaf before editing; every selected path existed and was tracked. The formatter changed 323 files; the remaining 12 selected paths produced no Git diff after the explicit invocation and remain covered by the clean per-file check.

## Exact formatter decisions

Every selected file uses the nearest owning `Cargo.toml` below. All selected manifests inherit or resolve to the edition shown. No closer rustfmt configuration exists in the selected source trees; historical docs/performance snapshots elsewhere under the leaf are not selected. Every invocation therefore uses the root `rustfmt.toml` (SHA-256 `5b08a5db07e6acaee62319a09314b9302c8cbc0f8e6b88b5ee8695f6303383d3`) with `skip_children=true`; `style_edition = "2024"` and `reorder_imports = true` are therefore applied consistently.

| owning manifest | edition | selected paths |
| --- | ---: | ---: |
| `crates/litchi-docx/Cargo.toml` | `2024` | 18 |
| `crates/litchi-imgconv/Cargo.toml` | `2024` | 2 |
| `crates/litchi-iwa/Cargo.toml` | `2024` | 4 |
| `crates/litchi-markdown/Cargo.toml` | `2024` | 3 |
| `crates/litchi-odf-common/Cargo.toml` | `2024` | 5 |
| `crates/litchi-ooxml-common/Cargo.toml` | `2024` | 2 |
| `crates/litchi-opc/Cargo.toml` | `2024` | 2 |
| `crates/litchi-ops/Cargo.toml` | `2024` | 7 |
| `crates/litchi-ppt/Cargo.toml` | `2024` | 1 |
| `crates/litchi-pptx/Cargo.toml` | `2024` | 2 |
| `crates/litchi-rtf/Cargo.toml` | `2024` | 163 |
| `crates/litchi-xls/Cargo.toml` | `2024` | 8 |
| `crates/litchi-xlsx/Cargo.toml` | `2024` | 1 |
| `crates/litchi/Cargo.toml` | `2024` | 2 |
| `crates/soapberry-zip/Cargo.toml` | `2021` | 3 |
| `vendor/allsorts/Cargo.toml` | `2021` | 64 |
| `vendor/flate2/Cargo.toml` | `2018` | 39 |
| `vendor/zlib-rs/Cargo.toml` | `2021` | 9 |

The exact path-to-manifest/edition mapping follows; paths are relative to the leaf.

### `crates/litchi-docx/Cargo.toml` — edition `2024` (18 paths)

- `crates/litchi-docx/src/document/owned/threads.rs`
- `crates/litchi-docx/src/document/package/package.rs`
- `crates/litchi-docx/src/mail_merge/codec.rs`
- `crates/litchi-docx/src/mail_merge/mod.rs`
- `crates/litchi-docx/src/mail_merge/package.rs`
- `crates/litchi-docx/src/mail_merge/tests.rs`
- `crates/litchi-docx/src/mail_merge/transaction.rs`
- `crates/litchi-docx/src/modern_comments/codec.rs`
- `crates/litchi-docx/src/modern_comments/mod.rs`
- `crates/litchi-docx/src/modern_comments/package.rs`
- `crates/litchi-docx/src/package/package/merge.rs`
- `crates/litchi-docx/src/package/package/settings.rs`
- `crates/litchi-docx/src/settings/document/codec.rs`
- `crates/litchi-docx/src/settings/document/tests.rs`
- `crates/litchi-docx/src/source_backed/paragraph_copy_record.rs`
- `crates/litchi-docx/src/source_backed.rs`
- `crates/litchi-docx/src/validation.rs`
- `crates/litchi-docx/tests/validation.rs`

### `crates/litchi-imgconv/Cargo.toml` — edition `2024` (2 paths)

- `crates/litchi-imgconv/src/emf/parser.rs`
- `crates/litchi-imgconv/src/wmf/parser.rs`

### `crates/litchi-iwa/Cargo.toml` — edition `2024` (4 paths)

- `crates/litchi-iwa/src/keynote/mod.rs`
- `crates/litchi-iwa/src/numbers/mod.rs`
- `crates/litchi-iwa/src/pages/editor/images.rs`
- `crates/litchi-iwa/src/pages/mod.rs`

### `crates/litchi-markdown/Cargo.toml` — edition `2024` (3 paths)

- `crates/litchi-markdown/tests/conformance.rs`
- `crates/litchi-markdown/tests/normative.rs`
- `crates/litchi-markdown/tests/reader.rs`

### `crates/litchi-odf-common/Cargo.toml` — edition `2024` (5 paths)

- `crates/litchi-odf-common/src/core/package.rs`
- `crates/litchi-odf-common/src/core/source_publication/insertion.rs`
- `crates/litchi-odf-common/src/detect.rs`
- `crates/litchi-odf-common/src/error.rs`
- `crates/litchi-odf-common/src/package/codec.rs`

### `crates/litchi-ooxml-common/Cargo.toml` — edition `2024` (2 paths)

- `crates/litchi-ooxml-common/src/custom_xml/codec.rs`
- `crates/litchi-ooxml-common/src/custom_xml/tests.rs`

### `crates/litchi-opc/Cargo.toml` — edition `2024` (2 paths)

- `crates/litchi-opc/src/error.rs`
- `crates/litchi-opc/src/source_backed.rs`

### `crates/litchi-ops/Cargo.toml` — edition `2024` (7 paths)

- `crates/litchi-ops/src/error.rs`
- `crates/litchi-ops/src/format/rtf/tests.rs`
- `crates/litchi-ops/src/tests/document.rs`
- `crates/litchi-ops/src/tests/signature.rs`
- `crates/litchi-ops/tests/docx_transactions.rs`
- `crates/litchi-ops/tests/docx_unified_save.rs`
- `crates/litchi-ops/tests/xlsb_p0_writes.rs`

### `crates/litchi-ppt/Cargo.toml` — edition `2024` (1 paths)

- `crates/litchi-ppt/src/html_publish.rs`

### `crates/litchi-pptx/Cargo.toml` — edition `2024` (2 paths)

- `crates/litchi-pptx/src/validation.rs`
- `crates/litchi-pptx/tests/validation.rs`

### `crates/litchi-rtf/Cargo.toml` — edition `2024` (163 paths)

- `crates/litchi-rtf/src/api/facade.rs`
- `crates/litchi-rtf/src/codec/writer/tests.rs`
- `crates/litchi-rtf/src/content/field/tests.rs`
- `crates/litchi-rtf/src/edit.rs`
- `crates/litchi-rtf/src/lib.rs`
- `crates/litchi-rtf/src/model/document/model.rs`
- `crates/litchi-rtf/src/model/document/tests.rs`
- `crates/litchi-rtf/src/streaming.rs`
- `crates/litchi-rtf/tests/annotations.rs`
- `crates/litchi-rtf/tests/associated_character_formatting.rs`
- `crates/litchi-rtf/tests/barcode_display_fields.rs`
- `crates/litchi-rtf/tests/bidirectional.rs`
- `crates/litchi-rtf/tests/body_story_event_order.rs`
- `crates/litchi-rtf/tests/character_background_color.rs`
- `crates/litchi-rtf/tests/character_decorations.rs`
- `crates/litchi-rtf/tests/character_positioning.rs`
- `crates/litchi-rtf/tests/character_style_references.rs`
- `crates/litchi-rtf/tests/character_type_grid.rs`
- `crates/litchi-rtf/tests/column_breaks.rs`
- `crates/litchi-rtf/tests/custom_xml.rs`
- `crates/litchi-rtf/tests/data_store.rs`
- `crates/litchi-rtf/tests/default_tab_width.rs`
- `crates/litchi-rtf/tests/document_asian_grid_compatibility.rs`
- `crates/litchi-rtf/tests/document_background.rs`
- `crates/litchi-rtf/tests/document_booklet_printing.rs`
- `crates/litchi-rtf/tests/document_compatibility_policy.rs`
- `crates/litchi-rtf/tests/document_default_formatting.rs`
- `crates/litchi-rtf/tests/document_drawing_grid.rs`
- `crates/litchi-rtf/tests/document_east_asian_compatibility.rs`
- `crates/litchi-rtf/tests/document_embedding_policies.rs`
- `crates/litchi-rtf/tests/document_facing_page_layout.rs`
- `crates/litchi-rtf/tests/document_file_settings.rs`
- `crates/litchi-rtf/tests/document_info.rs`
- `crates/litchi-rtf/tests/document_layout_theme_settings.rs`
- `crates/litchi-rtf/tests/document_legacy_layout_compatibility.rs`
- `crates/litchi-rtf/tests/document_line_spacing_compatibility.rs`
- `crates/litchi-rtf/tests/document_origin.rs`
- `crates/litchi-rtf/tests/document_output_settings.rs`
- `crates/litchi-rtf/tests/document_privacy_policies.rs`
- `crates/litchi-rtf/tests/document_processing_settings.rs`
- `crates/litchi-rtf/tests/document_protection.rs`
- `crates/litchi-rtf/tests/document_rendering_settings.rs`
- `crates/litchi-rtf/tests/document_revision_policies.rs`
- `crates/litchi-rtf/tests/document_save_preferences.rs`
- `crates/litchi-rtf/tests/document_style_policies.rs`
- `crates/litchi-rtf/tests/document_style_restrictions.rs`
- `crates/litchi-rtf/tests/document_table_layout_compatibility.rs`
- `crates/litchi-rtf/tests/document_variables.rs`
- `crates/litchi-rtf/tests/document_view.rs`
- `crates/litchi-rtf/tests/document_word_2003_compatibility.rs`
- `crates/litchi-rtf/tests/document_xml_policies.rs`
- `crates/litchi-rtf/tests/drawing_story_ownership.rs`
- `crates/litchi-rtf/tests/drawing_substream_ownership.rs`
- `crates/litchi-rtf/tests/editable_regions.rs`
- `crates/litchi-rtf/tests/embedded_fonts.rs`
- `crates/litchi-rtf/tests/emphasis_marks.rs`
- `crates/litchi-rtf/tests/external_reference_redact.rs`
- `crates/litchi-rtf/tests/external_references.rs`
- `crates/litchi-rtf/tests/field_status.rs`
- `crates/litchi-rtf/tests/file_table.rs`
- `crates/litchi-rtf/tests/fit_text.rs`
- `crates/litchi-rtf/tests/floating_table_boundaries.rs`
- `crates/litchi-rtf/tests/floating_table_positioning.rs`
- `crates/litchi-rtf/tests/font_metadata.rs`
- `crates/litchi-rtf/tests/font_theme.rs`
- `crates/litchi-rtf/tests/form_field_properties.rs`
- `crates/litchi-rtf/tests/form_fields.rs`
- `crates/litchi-rtf/tests/generated_list_markers.rs`
- `crates/litchi-rtf/tests/generator.rs`
- `crates/litchi-rtf/tests/group_nesting_depth.rs`
- `crates/litchi-rtf/tests/hyphenation.rs`
- `crates/litchi-rtf/tests/kinsoku.rs`
- `crates/litchi-rtf/tests/languages.rs`
- `crates/litchi-rtf/tests/latent_styles.rs`
- `crates/litchi-rtf/tests/legacy_drawings.rs`
- `crates/litchi-rtf/tests/legacy_numbering.rs`
- `crates/litchi-rtf/tests/legacy_paragraph_numbering.rs`
- `crates/litchi-rtf/tests/legacy_text_boxes.rs`
- `crates/litchi-rtf/tests/leniency.rs`
- `crates/litchi-rtf/tests/list_level_metadata.rs`
- `crates/litchi-rtf/tests/list_pictures.rs`
- `crates/litchi-rtf/tests/lists.rs`
- `crates/litchi-rtf/tests/mail_merge.rs`
- `crates/litchi-rtf/tests/math_properties.rs`
- `crates/litchi-rtf/tests/math_zones.rs`
- `crates/litchi-rtf/tests/navigation_entries.rs`
- `crates/litchi-rtf/tests/nested_tables.rs`
- `crates/litchi-rtf/tests/note_options.rs`
- `crates/litchi-rtf/tests/note_separators.rs`
- `crates/litchi-rtf/tests/objects.rs`
- `crates/litchi-rtf/tests/opaque_preservation.rs`
- `crates/litchi-rtf/tests/page_borders.rs`
- `crates/litchi-rtf/tests/page_breaks.rs`
- `crates/litchi-rtf/tests/paragraph_borders.rs`
- `crates/litchi-rtf/tests/paragraph_drop_caps.rs`
- `crates/litchi-rtf/tests/paragraph_frames.rs`
- `crates/litchi-rtf/tests/paragraph_groups.rs`
- `crates/litchi-rtf/tests/paragraph_layout_transactions.rs`
- `crates/litchi-rtf/tests/paragraph_line_breaking.rs`
- `crates/litchi-rtf/tests/paragraph_logical_indentation.rs`
- `crates/litchi-rtf/tests/paragraph_shading.rs`
- `crates/litchi-rtf/tests/paragraph_spacing_policy.rs`
- `crates/litchi-rtf/tests/paragraph_style_references.rs`
- `crates/litchi-rtf/tests/paragraph_tabs.rs`
- `crates/litchi-rtf/tests/parse_limits.rs`
- `crates/litchi-rtf/tests/picture_compatibility.rs`
- `crates/litchi-rtf/tests/picture_identity.rs`
- `crates/litchi-rtf/tests/picture_layout_metadata.rs`
- `crates/litchi-rtf/tests/picture_properties.rs`
- `crates/litchi-rtf/tests/protection_ranges.rs`
- `crates/litchi-rtf/tests/protection_users.rs`
- `crates/litchi-rtf/tests/review_display.rs`
- `crates/litchi-rtf/tests/revision_save.rs`
- `crates/litchi-rtf/tests/revisions.rs`
- `crates/litchi-rtf/tests/robustness.rs`
- `crates/litchi-rtf/tests/root_shapes.rs`
- `crates/litchi-rtf/tests/rsid_markers.rs`
- `crates/litchi-rtf/tests/section_columns.rs`
- `crates/litchi-rtf/tests/section_document_grid.rs`
- `crates/litchi-rtf/tests/section_line_numbering.rs`
- `crates/litchi-rtf/tests/section_note_options.rs`
- `crates/litchi-rtf/tests/section_page_number_heading.rs`
- `crates/litchi-rtf/tests/section_rendering.rs`
- `crates/litchi-rtf/tests/section_style_references.rs`
- `crates/litchi-rtf/tests/shape_binary_values.rs`
- `crates/litchi-rtf/tests/shape_group_stories.rs`
- `crates/litchi-rtf/tests/shape_groups.rs`
- `crates/litchi-rtf/tests/shape_hyperlinks.rs`
- `crates/litchi-rtf/tests/shape_mutation.rs`
- `crates/litchi-rtf/tests/shape_results.rs`
- `crates/litchi-rtf/tests/shape_text_frames.rs`
- `crates/litchi-rtf/tests/shape_theme_values.rs`
- `crates/litchi-rtf/tests/smart_tags_move_bookmarks.rs`
- `crates/litchi-rtf/tests/soft_breaks.rs`
- `crates/litchi-rtf/tests/spec_shading_and_quotes.rs`
- `crates/litchi-rtf/tests/special_characters.rs`
- `crates/litchi-rtf/tests/story_font_codepages.rs`
- `crates/litchi-rtf/tests/structural_revisions.rs`
- `crates/litchi-rtf/tests/style_list_filter.rs`
- `crates/litchi-rtf/tests/style_sort_method.rs`
- `crates/litchi-rtf/tests/stylesheets.rs`
- `crates/litchi-rtf/tests/table_autoformat_banding.rs`
- `crates/litchi-rtf/tests/table_borders_shading.rs`
- `crates/litchi-rtf/tests/table_distances.rs`
- `crates/litchi-rtf/tests/table_geometry.rs`
- `crates/litchi-rtf/tests/table_invisible_widths.rs`
- `crates/litchi-rtf/tests/table_layout.rs`
- `crates/litchi-rtf/tests/table_merge.rs`
- `crates/litchi-rtf/tests/table_row_cell_defaults.rs`
- `crates/litchi-rtf/tests/table_story_navigation_revisions.rs`
- `crates/litchi-rtf/tests/table_style_conditional.rs`
- `crates/litchi-rtf/tests/table_style_references.rs`
- `crates/litchi-rtf/tests/theme.rs`
- `crates/litchi-rtf/tests/title_outline_notes.rs`
- `crates/litchi-rtf/tests/typed_hyperlinks_and_references.rs`
- `crates/litchi-rtf/tests/underline_styles.rs`
- `crates/litchi-rtf/tests/user_properties.rs`
- `crates/litchi-rtf/tests/validation_report.rs`
- `crates/litchi-rtf/tests/window_caption.rs`
- `crates/litchi-rtf/tests/write_reservations.rs`
- `crates/litchi-rtf/tests/writer_text_delimiters.rs`
- `crates/litchi-rtf/tests/xml_namespaces.rs`
- `crates/litchi-rtf/tests/xsl_transform.rs`

### `crates/litchi-xls/Cargo.toml` — edition `2024` (8 paths)

- `crates/litchi-xls/src/cell_values/structural.rs`
- `crates/litchi-xls/src/package.rs`
- `crates/litchi-xls/src/sheet_features/mod.rs`
- `crates/litchi-xls/src/sheet_features/read.rs`
- `crates/litchi-xls/src/validation.rs`
- `crates/litchi-xls/src/workbook/codec/semantic/worksheet.rs`
- `crates/litchi-xls/src/writer/biff/cells.rs`
- `crates/litchi-xls/tests/validation.rs`

### `crates/litchi-xlsx/Cargo.toml` — edition `2024` (1 paths)

- `crates/litchi-xlsx/tests/encryption.rs`

### `crates/litchi/Cargo.toml` — edition `2024` (2 paths)

- `crates/litchi/examples/rtf_comprehensive_test.rs`
- `crates/litchi/examples/rtf_new_features_writer.rs`

### `crates/soapberry-zip/Cargo.toml` — edition `2021` (3 paths)

- `crates/soapberry-zip/src/office.rs`
- `crates/soapberry-zip/src/preserve.rs`
- `crates/soapberry-zip/tests/streaming_zip64_deflate.rs`

### `vendor/allsorts/Cargo.toml` — edition `2021` (64 paths)

- `vendor/allsorts/src/big5.rs`
- `vendor/allsorts/src/binary/budget.rs`
- `vendor/allsorts/src/binary/read.rs`
- `vendor/allsorts/src/binary/write.rs`
- `vendor/allsorts/src/binary.rs`
- `vendor/allsorts/src/bitmap/cbdt.rs`
- `vendor/allsorts/src/bitmap/sbix.rs`
- `vendor/allsorts/src/cff/cff2.rs`
- `vendor/allsorts/src/cff/charstring.rs`
- `vendor/allsorts/src/cff/outline.rs`
- `vendor/allsorts/src/cff/subset.rs`
- `vendor/allsorts/src/cff.rs`
- `vendor/allsorts/src/checksum.rs`
- `vendor/allsorts/src/context.rs`
- `vendor/allsorts/src/error.rs`
- `vendor/allsorts/src/font.rs`
- `vendor/allsorts/src/font_data.rs`
- `vendor/allsorts/src/font_specimen/scripts.rs`
- `vendor/allsorts/src/font_specimen.rs`
- `vendor/allsorts/src/glyph_info.rs`
- `vendor/allsorts/src/glyph_position.rs`
- `vendor/allsorts/src/gpos.rs`
- `vendor/allsorts/src/gsub.rs`
- `vendor/allsorts/src/layout/morx.rs`
- `vendor/allsorts/src/layout.rs`
- `vendor/allsorts/src/lib.rs`
- `vendor/allsorts/src/outline.rs`
- `vendor/allsorts/src/post.rs`
- `vendor/allsorts/src/scripts/arabic.rs`
- `vendor/allsorts/src/scripts/indic.rs`
- `vendor/allsorts/src/scripts/khmer.rs`
- `vendor/allsorts/src/scripts/mod.rs`
- `vendor/allsorts/src/scripts/mongolian.rs`
- `vendor/allsorts/src/scripts/myanmar.rs`
- `vendor/allsorts/src/scripts/syriac.rs`
- `vendor/allsorts/src/scripts/thai_lao.rs`
- `vendor/allsorts/src/scripts/tibetan.rs`
- `vendor/allsorts/src/subset.rs`
- `vendor/allsorts/src/tables/cmap/subset.rs`
- `vendor/allsorts/src/tables/cmap.rs`
- `vendor/allsorts/src/tables/colr.rs`
- `vendor/allsorts/src/tables/cpal.rs`
- `vendor/allsorts/src/tables/glyf/outline.rs`
- `vendor/allsorts/src/tables/glyf/subset.rs`
- `vendor/allsorts/src/tables/glyf/variation.rs`
- `vendor/allsorts/src/tables/glyf.rs`
- `vendor/allsorts/src/tables/kern.rs`
- `vendor/allsorts/src/tables/loca.rs`
- `vendor/allsorts/src/tables/morx.rs`
- `vendor/allsorts/src/tables/os2.rs`
- `vendor/allsorts/src/tables/svg.rs`
- `vendor/allsorts/src/tables/variable_fonts/avar.rs`
- `vendor/allsorts/src/tables/variable_fonts/cvar.rs`
- `vendor/allsorts/src/tables/variable_fonts/fvar.rs`
- `vendor/allsorts/src/tables/variable_fonts/gvar.rs`
- `vendor/allsorts/src/tables/variable_fonts/hvar.rs`
- `vendor/allsorts/src/tables/variable_fonts/stat.rs`
- `vendor/allsorts/src/tables/variable_fonts.rs`
- `vendor/allsorts/src/tables.rs`
- `vendor/allsorts/src/tests.rs`
- `vendor/allsorts/src/unicode/codepoint.rs`
- `vendor/allsorts/src/variations.rs`
- `vendor/allsorts/src/woff.rs`
- `vendor/allsorts/src/woff2.rs`

### `vendor/flate2/Cargo.toml` — edition `2018` (39 paths)

- `vendor/flate2/examples/compress_file.rs`
- `vendor/flate2/examples/decompress_file.rs`
- `vendor/flate2/examples/deflatedecoder-bufread.rs`
- `vendor/flate2/examples/deflatedecoder-read.rs`
- `vendor/flate2/examples/deflatedecoder-write.rs`
- `vendor/flate2/examples/deflateencoder-bufread.rs`
- `vendor/flate2/examples/deflateencoder-read.rs`
- `vendor/flate2/examples/deflateencoder-write.rs`
- `vendor/flate2/examples/gzdecoder-bufread.rs`
- `vendor/flate2/examples/gzdecoder-read.rs`
- `vendor/flate2/examples/gzdecoder-write.rs`
- `vendor/flate2/examples/gzencoder-bufread.rs`
- `vendor/flate2/examples/gzencoder-read.rs`
- `vendor/flate2/examples/gzencoder-write.rs`
- `vendor/flate2/examples/gzmultidecoder-bufread.rs`
- `vendor/flate2/examples/gzmultidecoder-read.rs`
- `vendor/flate2/examples/zlibdecoder-bufread.rs`
- `vendor/flate2/examples/zlibdecoder-read.rs`
- `vendor/flate2/examples/zlibdecoder-write.rs`
- `vendor/flate2/examples/zlibencoder-bufread.rs`
- `vendor/flate2/examples/zlibencoder-read.rs`
- `vendor/flate2/examples/zlibencoder-write.rs`
- `vendor/flate2/src/deflate/bufread.rs`
- `vendor/flate2/src/deflate/mod.rs`
- `vendor/flate2/src/ffi/c.rs`
- `vendor/flate2/src/ffi/miniz_oxide.rs`
- `vendor/flate2/src/ffi/mod.rs`
- `vendor/flate2/src/ffi/zlib_rs.rs`
- `vendor/flate2/src/gz/bufread.rs`
- `vendor/flate2/src/gz/mod.rs`
- `vendor/flate2/src/gz/read.rs`
- `vendor/flate2/src/gz/write.rs`
- `vendor/flate2/src/mem.rs`
- `vendor/flate2/src/zio.rs`
- `vendor/flate2/src/zlib/bufread.rs`
- `vendor/flate2/src/zlib/mod.rs`
- `vendor/flate2/src/zlib/read.rs`
- `vendor/flate2/tests/capabilities.rs`
- `vendor/flate2/tests/gunzip.rs`

### `vendor/zlib-rs/Cargo.toml` — edition `2021` (9 paths)

- `vendor/zlib-rs/src/adler32/avx2.rs`
- `vendor/zlib-rs/src/adler32/avx512.rs`
- `vendor/zlib-rs/src/adler32/avx512_vnni.rs`
- `vendor/zlib-rs/src/adler32/lsx.rs`
- `vendor/zlib-rs/src/adler32/neon.rs`
- `vendor/zlib-rs/src/cpu_features.rs`
- `vendor/zlib-rs/src/crc32/pclmulqdq.rs`
- `vendor/zlib-rs/src/crc32/vpclmulqdq.rs`
- `vendor/zlib-rs/src/stable.rs`

## Source-only command and review

For each edition group I ran the direct formatter with the explicit file list:

```text
rustfmt --edition <manifest-edition> --config-path rustfmt.toml --config skip_children=true <selected files>
```

After formatting, the same 335 explicit files pass the corresponding `rustfmt --check` invocation with zero nonzero results, and `git diff --check` is clean. A lexical review found no changed comment or string-literal contents and no identifier or literal token changes; punctuation differences are limited to rustfmt's equivalent closure-brace elision and trailing match-arm punctuation. No API, lint, test behavior, fixture, or configuration change was made. No Cargo, cargo fmt, rustdoc, build, glossary, native, corpus or performance command was run.

The final actual workspace formatter pass remains required after all source merges; this scoped repair does not claim that gate.
