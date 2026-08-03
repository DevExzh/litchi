//! Compatibility reference extraction for legacy IWA payloads.

use crate::Result;
use crate::archive::ArchiveObject;
use crate::ref_graph::{ObjectId, ReferenceGraph};

pub(super) fn extract(
    source_id: ObjectId,
    object: &ArchiveObject,
    graph: &mut ReferenceGraph,
) -> Result<()> {
    use prost::Message;

    // For each raw message, try to extract references
    for raw_msg in &object.messages {
        let msg_type = raw_msg.type_;

        // Extract references based on message type
        // We decode the specific protobuf message and extract its Reference fields
        match msg_type {
            // TST (Table) types
            6000 | 6001 => {
                // TST.TableModelArchive contains multiple style and data references
                if let Ok(table) = crate::protobuf::tst::TableModelArchive::decode(&*raw_msg.data) {
                    // Extract style references
                    extract_reference(source_id, graph, &table.table_style);
                    extract_reference(source_id, graph, &table.body_text_style);
                    extract_reference(source_id, graph, &table.header_row_text_style);
                    extract_reference(source_id, graph, &table.header_column_text_style);
                    extract_reference(source_id, graph, &table.footer_row_text_style);
                    extract_reference(source_id, graph, &table.body_cell_style);
                    extract_reference(source_id, graph, &table.header_row_style);
                    extract_reference(source_id, graph, &table.header_column_style);
                    extract_reference(source_id, graph, &table.footer_row_style);

                    // Extract optional style references
                    if let Some(ref table_name_style) = table.table_name_style {
                        extract_reference(source_id, graph, table_name_style);
                    }
                    if let Some(ref table_name_shape_style) = table.table_name_shape_style {
                        extract_reference(source_id, graph, table_name_shape_style);
                    }

                    // Extract data store sub-references
                    // DataStore contains references to column_headers, string_table, style_table, etc.
                    let data_store = &table.base_data_store;
                    extract_reference(source_id, graph, &data_store.column_headers);
                    extract_reference(source_id, graph, &data_store.string_table);
                    extract_reference(source_id, graph, &data_store.style_table);
                    extract_reference(source_id, graph, &data_store.formula_table);
                    extract_reference(source_id, graph, &data_store.format_table_pre_bnc);
                    if let Some(format_table) = &data_store.format_table {
                        extract_reference(source_id, graph, format_table);
                    }

                    // Optional references
                    if let Some(ref formula_error_table) = data_store.formula_error_table {
                        extract_reference(source_id, graph, formula_error_table);
                    }
                    if let Some(ref choice_list) = data_store.multiple_choice_list_format_table {
                        extract_reference(source_id, graph, choice_list);
                    }
                    if let Some(ref merge_map) = data_store.merge_region_map {
                        extract_reference(source_id, graph, merge_map);
                    }
                }
            },

            6005 | 6201 => {
                if let Ok(list) = crate::protobuf::tst::TableDataList::decode(&*raw_msg.data) {
                    for segment in &list.segments {
                        extract_reference(source_id, graph, segment);
                    }
                    for entry in &list.entries {
                        extract_table_data_list_entry_references(graph, source_id, entry);
                    }
                }
            },

            6011 => {
                if let Ok(segment) =
                    crate::protobuf::tst::TableDataListSegment::decode(&*raw_msg.data)
                {
                    for entry in &segment.entries {
                        extract_table_data_list_entry_references(graph, source_id, entry);
                    }
                }
            },

            // TSWP (Word Processing/Text) types
            2001..=2022 => {
                // TSWP.StorageArchive contains text content and may reference styles
                if let Ok(storage) = crate::protobuf::tswp::StorageArchive::decode(&*raw_msg.data) {
                    // Extract stylesheet reference if present
                    if let Some(ref style_sheet) = storage.style_sheet {
                        extract_reference(source_id, graph, style_sheet);
                    }

                    // Note: Attachments are stored in separate fields in the attribute tables
                    // They're not directly accessible as simple references in StorageArchive
                }
            },

            // KN (Keynote) types
            5 | 6 => {
                // KN.SlideArchive contains references to drawables, builds, and transitions
                if let Ok(slide) = crate::protobuf::kn::SlideArchive::decode(&*raw_msg.data) {
                    // Extract style reference
                    extract_reference(source_id, graph, &slide.style);

                    // Extract drawable references (shapes, images, text boxes)
                    for drawable in &slide.owned_drawables {
                        extract_reference(source_id, graph, drawable);
                    }

                    // Extract build animation references
                    for build in &slide.builds {
                        extract_reference(source_id, graph, build);
                    }

                    // Extract placeholder references
                    if let Some(ref title) = slide.title_placeholder {
                        extract_reference(source_id, graph, title);
                    }
                    if let Some(ref body) = slide.body_placeholder {
                        extract_reference(source_id, graph, body);
                    }
                    if let Some(ref object) = slide.object_placeholder {
                        extract_reference(source_id, graph, object);
                    }
                    if let Some(ref slide_num) = slide.slide_number_placeholder {
                        extract_reference(source_id, graph, slide_num);
                    }

                    // Extract style references
                    for para_style in &slide.body_paragraph_styles {
                        extract_reference(source_id, graph, para_style);
                    }
                    for list_style in &slide.body_list_styles {
                        extract_reference(source_id, graph, list_style);
                    }
                }
            },

            2 => {
                // KN.ShowArchive (conflicts with TSP.MessageInfo, handle by context)
                // Try to decode as ShowArchive for Keynote documents
                if let Ok(show) = crate::protobuf::kn::ShowArchive::decode(&*raw_msg.data) {
                    // Extract theme and stylesheet references
                    extract_reference(source_id, graph, &show.theme);
                    extract_reference(source_id, graph, &show.stylesheet);

                    // Extract UI state reference
                    if let Some(ref ui_state) = show.ui_state {
                        extract_reference(source_id, graph, ui_state);
                    }

                    // Extract recording reference if present
                    if let Some(ref recording) = show.recording {
                        extract_reference(source_id, graph, recording);
                    }

                    // Note: Slide references are in the slide_tree structure
                    // which is not a simple Reference type
                }
            },

            // TN (Numbers) types
            3 => {
                // TN.SheetArchive / TN.FormBasedSheetArchive
                if let Ok(sheet) = crate::protobuf::tn::SheetArchive::decode(&*raw_msg.data) {
                    // Extract drawable info references
                    for drawable_ref in &sheet.drawable_infos {
                        extract_reference(source_id, graph, drawable_ref);
                    }

                    for header in &sheet.headers {
                        extract_reference(source_id, graph, header);
                    }
                    for footer in &sheet.footers {
                        extract_reference(source_id, graph, footer);
                    }

                    // Old documents used one storage reference for each area.
                    if sheet.headers.is_empty() && sheet.footers.is_empty() {
                        extract_legacy_sheet_headers(source_id, graph, &sheet);
                    }
                }
            },

            // TSD (Drawing/Shape) types
            // Implementation Status: ✓ COMPLETED (2025-11-04)
            // Based on TSDArchives.proto and libetonyek's reference extraction
            3002 => {
                // TSD.DrawableArchive - base type for all drawables
                if let Ok(drawable) = crate::protobuf::tsd::DrawableArchive::decode(&*raw_msg.data)
                {
                    // Extract parent reference (drawable hierarchy)
                    if let Some(ref parent) = drawable.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    // Note: geometry is not a reference, just position/size data
                    // exterior_text_wrap is configuration, not a reference
                }
            },
            3003 => {
                // TSD.ContainerArchive - container for grouped objects
                if let Ok(container) =
                    crate::protobuf::tsd::ContainerArchive::decode(&*raw_msg.data)
                {
                    // Extract parent reference
                    if let Some(ref parent) = container.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    // Extract all child references
                    for child in &container.children {
                        extract_reference(source_id, graph, child);
                    }
                }
            },
            3004 => {
                // TSD.ShapeArchive - shapes (rectangles, circles, polygons, etc.)
                if let Ok(shape) = crate::protobuf::tsd::ShapeArchive::decode(&*raw_msg.data) {
                    // ShapeArchive embeds DrawableArchive in 'super' field (required)
                    // Extract parent from the super DrawableArchive
                    if let Some(ref parent) = shape.super_.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    // Extract style reference
                    if let Some(ref style) = shape.style {
                        extract_reference(source_id, graph, style);
                    }
                    // Note: pathsource, head_line_end, tail_line_end are not references
                    // but embedded data structures
                }
            },
            3005 => {
                // TSD.ImageArchive - images
                if let Ok(image) = crate::protobuf::tsd::ImageArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = image.super_.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    // Extract style reference
                    if let Some(ref style) = image.style {
                        extract_reference(source_id, graph, style);
                    }
                    // Note: data field is a DataReference, not an object Reference
                    // database_originalData is also for media assets
                }
            },
            3006 => {
                // TSD.MaskArchive - image masks
                if let Ok(mask) = crate::protobuf::tsd::MaskArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = mask.super_.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    // Note: pathsource is embedded data, not a reference
                }
            },
            3007 => {
                // TSD.MovieArchive - video objects
                if let Ok(movie) = crate::protobuf::tsd::MovieArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = movie.super_.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    // Extract style reference
                    if let Some(ref style) = movie.style {
                        extract_reference(source_id, graph, style);
                    }
                    // Note: movieData is a DataReference, not an object Reference
                }
            },
            3008 => {
                // TSD.GroupArchive - grouped shapes/objects
                if let Ok(group) = crate::protobuf::tsd::GroupArchive::decode(&*raw_msg.data) {
                    // Extract parent from super DrawableArchive (required field)
                    if let Some(ref parent) = group.super_.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    // Extract all child references (objects in the group)
                    for child in &group.children {
                        extract_reference(source_id, graph, child);
                    }
                }
            },
            3009 => {
                // TSD.ConnectionLineArchive - connector lines between shapes
                if let Ok(conn_line) =
                    crate::protobuf::tsd::ConnectionLineArchive::decode(&*raw_msg.data)
                {
                    // Extract parent and style from super ShapeArchive (required field)
                    // ConnectionLineArchive.super_ is ShapeArchive
                    // ShapeArchive.super_ is DrawableArchive
                    if let Some(ref parent) = conn_line.super_.super_.parent {
                        extract_reference(source_id, graph, parent);
                    }
                    if let Some(ref style) = conn_line.super_.style {
                        extract_reference(source_id, graph, style);
                    }
                    // Extract connection endpoints
                    if let Some(ref connected_from) = conn_line.connected_from {
                        extract_reference(source_id, graph, connected_from);
                    }
                    if let Some(ref connected_to) = conn_line.connected_to {
                        extract_reference(source_id, graph, connected_to);
                    }
                }
            },
            3056 => {
                if let Ok(comment) =
                    crate::protobuf::tsd::CommentStorageArchive::decode(&*raw_msg.data)
                {
                    if let Some(author) = &comment.author {
                        extract_reference(source_id, graph, author);
                    }
                    for reply in &comment.replies {
                        extract_reference(source_id, graph, reply);
                    }
                }
            },

            // TSCH (Chart) types
            // Implementation Status: ✓ COMPLETED (2025-11-04)
            // Based on TSCHArchives.proto and libetonyek's chart parsing
            5000 => {
                // TSCH.PreUFF.ChartInfoArchive - legacy chart format
                // This is a pre-unified format chart, structure may vary
                // Attempt basic reference extraction but may fail gracefully
                if let Ok(chart_info) =
                    crate::protobuf::tsch::pre_uff::ChartInfoArchive::decode(&*raw_msg.data)
                {
                    // Extract chart style reference if present
                    if let Some(ref style) = chart_info.style {
                        extract_reference(source_id, graph, style);
                    }
                    // Note: PreUFF ChartInfoArchive doesn't have a direct legend field
                    // Legend info is embedded in other structures
                }
            },
            5004 => {
                // TSCH.ChartMediatorArchive - mediator between chart and data
                if let Ok(mediator) =
                    crate::protobuf::tsch::ChartMediatorArchive::decode(&*raw_msg.data)
                {
                    // Extract info reference (points to the chart drawable)
                    if let Some(ref info) = mediator.info {
                        extract_reference(source_id, graph, info);
                    }
                    // Note: local_series_indexes and remote_series_indexes are
                    // indices, not references to objects
                }
            },
            5020 => {
                // TSCH.ChartStylePreset - preset styles for charts
                if let Ok(preset) = crate::protobuf::tsch::ChartStylePreset::decode(&*raw_msg.data)
                {
                    // Extract chart style reference
                    if let Some(ref chart_style) = preset.chart_style {
                        extract_reference(source_id, graph, chart_style);
                    }
                    // Extract legend style reference
                    if let Some(ref legend_style) = preset.legend_style {
                        extract_reference(source_id, graph, legend_style);
                    }
                    // Note: ChartStylePreset has a complex nested structure
                    // Styles for series and axes are managed through different fields
                    // than what might be expected from the pre-UFF format
                }
            },
            5021 => {
                // TSCH.ChartDrawableArchive - main chart drawable
                if let Ok(chart_drawable) =
                    crate::protobuf::tsch::ChartDrawableArchive::decode(&*raw_msg.data)
                {
                    // Extract parent from super DrawableArchive
                    if let Some(ref drawable) = chart_drawable.super_
                        && let Some(ref parent) = drawable.parent
                    {
                        extract_reference(source_id, graph, parent);
                    }
                    // Note: ChartArchive is embedded via protobuf extensions,
                    // which requires special handling. The chart data and preset
                    // references would be in the extension fields that we can't
                    // easily access through the standard decode.
                }
            },

            // TP (Pages) types
            10000 => {
                // TP.DocumentArchive
                if let Ok(doc) = crate::protobuf::tp::DocumentArchive::decode(&*raw_msg.data) {
                    for reference in [
                        doc.stylesheet.as_ref(),
                        doc.floating_drawables.as_ref(),
                        doc.body_storage.as_ref(),
                        doc.section.as_ref(),
                        doc.theme.as_ref(),
                        doc.settings.as_ref(),
                        doc.deprecated_layout_state.as_ref(),
                        doc.deprecated_view_state.as_ref(),
                        doc.most_recent_change_session.as_ref(),
                        doc.drawables_zorder.as_ref(),
                        doc.tables_custom_format_list.as_ref(),
                        doc.flow_info_container.as_ref(),
                        doc.merge_data.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, graph, reference);
                    }
                    for reference in doc
                        .citation_records
                        .iter()
                        .chain(&doc.toc_styles)
                        .chain(&doc.change_sessions)
                        .chain(&doc.page_templates)
                    {
                        extract_reference(source_id, graph, reference);
                    }
                    let tsa = &doc.super_;
                    for reference in [
                        tsa.calculation_engine.as_ref(),
                        tsa.view_state.as_ref(),
                        tsa.function_browser_state.as_ref(),
                        tsa.tables_custom_format_list.as_ref(),
                        tsa.shortcut_controller.as_ref(),
                        tsa.annotation_cache_deprecated.as_ref(),
                        tsa.custom_format_list.as_ref(),
                        tsa.annotation_cache_deprecated_2.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, graph, reference);
                    }
                    let tsk = &tsa.super_;
                    for reference in [
                        tsk.annotation_author_storage.as_ref(),
                        tsk.collaboration_operation_history.as_ref(),
                        tsk.activity_stream.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, graph, reference);
                    }
                    for reference in &tsk.activity_log_entries {
                        extract_reference(source_id, graph, reference);
                    }
                }
            },

            10011 => {
                if let Ok(section) = crate::protobuf::tp::SectionArchive::decode(&*raw_msg.data) {
                    for reference in section
                        .obsolete_headers
                        .iter()
                        .chain(&section.obsolete_footers)
                        .chain(&section.obsolete_section_template_drawables)
                    {
                        extract_reference(source_id, graph, reference);
                    }
                    for reference in [
                        section.first_section_template_page.as_ref(),
                        section.even_section_template_page.as_ref(),
                        section.odd_section_template_page.as_ref(),
                        section.user_defined_guide_storage.as_ref(),
                    ]
                    .into_iter()
                    .flatten()
                    {
                        extract_reference(source_id, graph, reference);
                    }
                }
            },

            10143 => {
                if let Ok(template) =
                    crate::protobuf::tp::SectionTemplateArchive::decode(&*raw_msg.data)
                {
                    for reference in template
                        .headers
                        .iter()
                        .chain(&template.footers)
                        .chain(&template.section_template_drawables)
                    {
                        extract_reference(source_id, graph, reference);
                    }
                }
            },

            _ => {
                // For unknown types, we don't extract references
                // This is fine as we handle the most common types above
            },
        }
    }

    Ok(())
}

fn extract_reference(
    source_id: ObjectId,
    graph: &mut ReferenceGraph,
    reference: &crate::protobuf::tsp::Reference,
) {
    if let Some(target_id) = ObjectId::new(reference.identifier) {
        graph.add_object_reference(source_id, target_id);
    }
}

#[allow(deprecated)]
fn extract_legacy_sheet_headers(
    source_id: ObjectId,
    graph: &mut ReferenceGraph,
    sheet: &crate::protobuf::tn::SheetArchive,
) {
    if let Some(header) = &sheet.header_storage {
        extract_reference(source_id, graph, header);
    }
    if let Some(footer) = &sheet.footer_storage {
        extract_reference(source_id, graph, footer);
    }
}

fn extract_table_data_list_entry_references(
    graph: &mut ReferenceGraph,
    source_id: ObjectId,
    entry: &crate::protobuf::tst::table_data_list::ListEntry,
) {
    for reference in [
        entry.reference.as_ref(),
        entry.rich_text_payload.as_ref(),
        entry.comment_storage.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        extract_reference(source_id, graph, reference);
    }
}
