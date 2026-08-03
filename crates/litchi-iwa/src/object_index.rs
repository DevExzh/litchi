//! Object Index for Cross-Referencing in iWork Documents
//!
//! iWork documents contain an object index that maps object IDs to their
//! locations in IWA files. This allows objects to reference each other
//! across different archive files.

use std::collections::HashMap;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::bundle::Bundle;
use crate::ref_graph::{ObjectId, ObjectIdIter, ReferenceGraph};
use crate::{Error, Result};

/// Represents an entry in the object index
#[derive(Debug, Clone)]
pub struct ObjectIndexEntry {
    /// Unique object identifier
    pub id: u64,
    /// Which IWA file contains this object
    pub fragment_name: String,
    /// Offset within the IWA file
    pub data_offset: u64,
    /// Length of the object data
    pub data_length: u64,
    /// Type of the object
    pub object_type: u32,
}

impl ObjectIndexEntry {
    /// Return the validated object identity, if this compatibility entry is
    /// non-null.
    pub fn object_id(&self) -> Option<ObjectId> {
        ObjectId::new(self.id)
    }
}

/// Object index that maps object IDs to their locations
#[derive(Debug, Clone)]
pub struct ObjectIndex {
    /// Map from object ID to index entry
    entries: HashMap<u64, ObjectIndexEntry>,
    /// Map from fragment name to list of object IDs
    fragment_objects: HashMap<String, Vec<u64>>,
    /// Reference graph tracking object dependencies
    reference_graph: ReferenceGraph,
}

impl Default for ObjectIndex {
    fn default() -> Self {
        Self::new()
    }
}

impl ObjectIndex {
    /// Create an empty object index
    pub fn new() -> Self {
        Self {
            entries: HashMap::new(),
            fragment_objects: HashMap::new(),
            reference_graph: ReferenceGraph::new(),
        }
    }

    /// Build object index from a bundle
    pub fn from_bundle(bundle: &Bundle) -> Result<Self> {
        let mut index = Self::new();

        // Parse all archives to build the index
        for (archive_name, archive) in bundle.archives() {
            index.parse_archive(archive_name, archive)?;
        }

        Ok(index)
    }

    /// Parse an archive to extract object information
    ///
    /// This extracts position information for each object in the archive,
    /// allowing for efficient lazy loading and partial parsing. The implementation
    /// follows the approach used by libetonyek's IWAObjectIndex.
    ///
    /// # Implementation Status
    ///
    /// ✓ COMPLETED: Proper data_offset and data_length calculation (2025-11-04)
    ///   - Tracks byte positions during archive parsing
    ///   - Enables efficient random access to objects
    ///   - Follows libetonyek's ObjectRecord approach
    fn parse_archive(&mut self, archive_name: &str, archive: &Archive) -> Result<()> {
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.ok_or_else(|| {
                Error::Archive(format!(
                    "archive {archive_name} contains an object without an identifier"
                ))
            })?;
            let object_id = ObjectId::try_from(identifier).map_err(|_| {
                Error::Archive(format!(
                    "archive {archive_name} contains the null object identifier"
                ))
            })?;

            // Determine object type from first message
            let object_type = object.messages.first().map(|msg| msg.type_).unwrap_or(0);

            let entry = ObjectIndexEntry {
                id: identifier,
                fragment_name: archive_name.to_string(),
                // Use actual byte offsets from the parsed archive
                // These match the approach used in libetonyek's ObjectRecord
                data_offset: object.data_offset,
                data_length: object.data_length,
                object_type,
            };

            self.entries.insert(identifier, entry);
            self.fragment_objects
                .entry(archive_name.to_string())
                .or_default()
                .push(identifier);

            // MessageInfo is the authoritative, application-independent
            // reference index emitted by iWork for every payload.
            let mut has_indexed_references = false;
            for message_info in &object.archive_info.message_infos {
                has_indexed_references |= !message_info.object_references.is_empty();
                for &reference in &message_info.object_references {
                    if let Some(target_id) = ObjectId::new(reference) {
                        self.reference_graph
                            .add_object_reference(object_id, target_id);
                    }
                }
            }

            // Some old archives omit MessageInfo references. Decode only
            // unambiguous high-numbered payloads as a compatibility fallback;
            // low message types overlap between Numbers and Keynote.
            if !has_indexed_references && object_type >= 2000 {
                self.parse_object_references(identifier, object)?;
            }
        }
        Ok(())
    }

    /// Parse object references within an object's messages
    ///
    /// This function extracts TSP.Reference fields from protobuf messages and builds
    /// the object reference graph. iWork documents use object references extensively
    /// to connect related objects (e.g., tables reference their data stores, slides
    /// reference their drawables, etc.).
    fn parse_object_references(&mut self, object_id: u64, object: &ArchiveObject) -> Result<()> {
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
                    if let Ok(table) =
                        crate::protobuf::tst::TableModelArchive::decode(&*raw_msg.data)
                    {
                        // Extract style references
                        self.extract_reference(object_id, &table.table_style);
                        self.extract_reference(object_id, &table.body_text_style);
                        self.extract_reference(object_id, &table.header_row_text_style);
                        self.extract_reference(object_id, &table.header_column_text_style);
                        self.extract_reference(object_id, &table.footer_row_text_style);
                        self.extract_reference(object_id, &table.body_cell_style);
                        self.extract_reference(object_id, &table.header_row_style);
                        self.extract_reference(object_id, &table.header_column_style);
                        self.extract_reference(object_id, &table.footer_row_style);

                        // Extract optional style references
                        if let Some(ref table_name_style) = table.table_name_style {
                            self.extract_reference(object_id, table_name_style);
                        }
                        if let Some(ref table_name_shape_style) = table.table_name_shape_style {
                            self.extract_reference(object_id, table_name_shape_style);
                        }

                        // Extract data store sub-references
                        // DataStore contains references to column_headers, string_table, style_table, etc.
                        let data_store = &table.base_data_store;
                        self.extract_reference(object_id, &data_store.column_headers);
                        self.extract_reference(object_id, &data_store.string_table);
                        self.extract_reference(object_id, &data_store.style_table);
                        self.extract_reference(object_id, &data_store.formula_table);
                        self.extract_reference(object_id, &data_store.format_table_pre_bnc);
                        if let Some(format_table) = &data_store.format_table {
                            self.extract_reference(object_id, format_table);
                        }

                        // Optional references
                        if let Some(ref formula_error_table) = data_store.formula_error_table {
                            self.extract_reference(object_id, formula_error_table);
                        }
                        if let Some(ref choice_list) = data_store.multiple_choice_list_format_table
                        {
                            self.extract_reference(object_id, choice_list);
                        }
                        if let Some(ref merge_map) = data_store.merge_region_map {
                            self.extract_reference(object_id, merge_map);
                        }
                    }
                },

                6005 | 6201 => {
                    if let Ok(list) = crate::protobuf::tst::TableDataList::decode(&*raw_msg.data) {
                        for segment in &list.segments {
                            self.extract_reference(object_id, segment);
                        }
                        for entry in &list.entries {
                            extract_table_data_list_entry_references(self, object_id, entry);
                        }
                    }
                },

                6011 => {
                    if let Ok(segment) =
                        crate::protobuf::tst::TableDataListSegment::decode(&*raw_msg.data)
                    {
                        for entry in &segment.entries {
                            extract_table_data_list_entry_references(self, object_id, entry);
                        }
                    }
                },

                // TSWP (Word Processing/Text) types
                2001..=2022 => {
                    // TSWP.StorageArchive contains text content and may reference styles
                    if let Ok(storage) =
                        crate::protobuf::tswp::StorageArchive::decode(&*raw_msg.data)
                    {
                        // Extract stylesheet reference if present
                        if let Some(ref style_sheet) = storage.style_sheet {
                            self.extract_reference(object_id, style_sheet);
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
                        self.extract_reference(object_id, &slide.style);

                        // Extract drawable references (shapes, images, text boxes)
                        for drawable in &slide.owned_drawables {
                            self.extract_reference(object_id, drawable);
                        }

                        // Extract build animation references
                        for build in &slide.builds {
                            self.extract_reference(object_id, build);
                        }

                        // Extract placeholder references
                        if let Some(ref title) = slide.title_placeholder {
                            self.extract_reference(object_id, title);
                        }
                        if let Some(ref body) = slide.body_placeholder {
                            self.extract_reference(object_id, body);
                        }
                        if let Some(ref object) = slide.object_placeholder {
                            self.extract_reference(object_id, object);
                        }
                        if let Some(ref slide_num) = slide.slide_number_placeholder {
                            self.extract_reference(object_id, slide_num);
                        }

                        // Extract style references
                        for para_style in &slide.body_paragraph_styles {
                            self.extract_reference(object_id, para_style);
                        }
                        for list_style in &slide.body_list_styles {
                            self.extract_reference(object_id, list_style);
                        }
                    }
                },

                2 => {
                    // KN.ShowArchive (conflicts with TSP.MessageInfo, handle by context)
                    // Try to decode as ShowArchive for Keynote documents
                    if let Ok(show) = crate::protobuf::kn::ShowArchive::decode(&*raw_msg.data) {
                        // Extract theme and stylesheet references
                        self.extract_reference(object_id, &show.theme);
                        self.extract_reference(object_id, &show.stylesheet);

                        // Extract UI state reference
                        if let Some(ref ui_state) = show.ui_state {
                            self.extract_reference(object_id, ui_state);
                        }

                        // Extract recording reference if present
                        if let Some(ref recording) = show.recording {
                            self.extract_reference(object_id, recording);
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
                            self.extract_reference(object_id, drawable_ref);
                        }

                        for header in &sheet.headers {
                            self.extract_reference(object_id, header);
                        }
                        for footer in &sheet.footers {
                            self.extract_reference(object_id, footer);
                        }

                        // Old documents used one storage reference for each area.
                        if sheet.headers.is_empty() && sheet.footers.is_empty() {
                            self.extract_legacy_sheet_headers(object_id, &sheet);
                        }
                    }
                },

                // TSD (Drawing/Shape) types
                // Implementation Status: ✓ COMPLETED (2025-11-04)
                // Based on TSDArchives.proto and libetonyek's reference extraction
                3002 => {
                    // TSD.DrawableArchive - base type for all drawables
                    if let Ok(drawable) =
                        crate::protobuf::tsd::DrawableArchive::decode(&*raw_msg.data)
                    {
                        // Extract parent reference (drawable hierarchy)
                        if let Some(ref parent) = drawable.parent {
                            self.extract_reference(object_id, parent);
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
                            self.extract_reference(object_id, parent);
                        }
                        // Extract all child references
                        for child in &container.children {
                            self.extract_reference(object_id, child);
                        }
                    }
                },
                3004 => {
                    // TSD.ShapeArchive - shapes (rectangles, circles, polygons, etc.)
                    if let Ok(shape) = crate::protobuf::tsd::ShapeArchive::decode(&*raw_msg.data) {
                        // ShapeArchive embeds DrawableArchive in 'super' field (required)
                        // Extract parent from the super DrawableArchive
                        if let Some(ref parent) = shape.super_.parent {
                            self.extract_reference(object_id, parent);
                        }
                        // Extract style reference
                        if let Some(ref style) = shape.style {
                            self.extract_reference(object_id, style);
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
                            self.extract_reference(object_id, parent);
                        }
                        // Extract style reference
                        if let Some(ref style) = image.style {
                            self.extract_reference(object_id, style);
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
                            self.extract_reference(object_id, parent);
                        }
                        // Note: pathsource is embedded data, not a reference
                    }
                },
                3007 => {
                    // TSD.MovieArchive - video objects
                    if let Ok(movie) = crate::protobuf::tsd::MovieArchive::decode(&*raw_msg.data) {
                        // Extract parent from super DrawableArchive (required field)
                        if let Some(ref parent) = movie.super_.parent {
                            self.extract_reference(object_id, parent);
                        }
                        // Extract style reference
                        if let Some(ref style) = movie.style {
                            self.extract_reference(object_id, style);
                        }
                        // Note: movieData is a DataReference, not an object Reference
                    }
                },
                3008 => {
                    // TSD.GroupArchive - grouped shapes/objects
                    if let Ok(group) = crate::protobuf::tsd::GroupArchive::decode(&*raw_msg.data) {
                        // Extract parent from super DrawableArchive (required field)
                        if let Some(ref parent) = group.super_.parent {
                            self.extract_reference(object_id, parent);
                        }
                        // Extract all child references (objects in the group)
                        for child in &group.children {
                            self.extract_reference(object_id, child);
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
                            self.extract_reference(object_id, parent);
                        }
                        if let Some(ref style) = conn_line.super_.style {
                            self.extract_reference(object_id, style);
                        }
                        // Extract connection endpoints
                        if let Some(ref connected_from) = conn_line.connected_from {
                            self.extract_reference(object_id, connected_from);
                        }
                        if let Some(ref connected_to) = conn_line.connected_to {
                            self.extract_reference(object_id, connected_to);
                        }
                    }
                },
                3056 => {
                    if let Ok(comment) =
                        crate::protobuf::tsd::CommentStorageArchive::decode(&*raw_msg.data)
                    {
                        if let Some(author) = &comment.author {
                            self.extract_reference(object_id, author);
                        }
                        for reply in &comment.replies {
                            self.extract_reference(object_id, reply);
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
                            self.extract_reference(object_id, style);
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
                            self.extract_reference(object_id, info);
                        }
                        // Note: local_series_indexes and remote_series_indexes are
                        // indices, not references to objects
                    }
                },
                5020 => {
                    // TSCH.ChartStylePreset - preset styles for charts
                    if let Ok(preset) =
                        crate::protobuf::tsch::ChartStylePreset::decode(&*raw_msg.data)
                    {
                        // Extract chart style reference
                        if let Some(ref chart_style) = preset.chart_style {
                            self.extract_reference(object_id, chart_style);
                        }
                        // Extract legend style reference
                        if let Some(ref legend_style) = preset.legend_style {
                            self.extract_reference(object_id, legend_style);
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
                            self.extract_reference(object_id, parent);
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
                            self.extract_reference(object_id, reference);
                        }
                        for reference in doc
                            .citation_records
                            .iter()
                            .chain(&doc.toc_styles)
                            .chain(&doc.change_sessions)
                            .chain(&doc.page_templates)
                        {
                            self.extract_reference(object_id, reference);
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
                            self.extract_reference(object_id, reference);
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
                            self.extract_reference(object_id, reference);
                        }
                        for reference in &tsk.activity_log_entries {
                            self.extract_reference(object_id, reference);
                        }
                    }
                },

                10011 => {
                    if let Ok(section) = crate::protobuf::tp::SectionArchive::decode(&*raw_msg.data)
                    {
                        for reference in section
                            .obsolete_headers
                            .iter()
                            .chain(&section.obsolete_footers)
                            .chain(&section.obsolete_section_template_drawables)
                        {
                            self.extract_reference(object_id, reference);
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
                            self.extract_reference(object_id, reference);
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
                            self.extract_reference(object_id, reference);
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

    /// Helper function to extract a single TSP.Reference
    ///
    /// Adds the referenced object ID to the reference graph, creating edges
    /// from source objects to their dependencies. This enables:
    /// - Dependency tracking (what objects does this reference?)
    /// - Reverse lookups (what objects reference this?)
    /// - Graph traversal for complete object resolution
    ///
    /// # Arguments
    ///
    /// * `source_id` - The object ID that contains this reference
    /// * `reference` - The TSP.Reference to extract and track
    ///
    /// # Performance
    ///
    /// O(1) average case for HashMap insertion. Uses efficient deduplication
    /// to avoid storing duplicate references.
    fn extract_reference(&mut self, source_id: u64, reference: &crate::protobuf::tsp::Reference) {
        let (Some(source_id), Some(target_id)) = (
            ObjectId::new(source_id),
            ObjectId::new(reference.identifier),
        ) else {
            return;
        };

        // Build the reference graph through the checked identity boundary.
        self.reference_graph
            .add_object_reference(source_id, target_id);
    }

    /// Get an object entry by ID
    pub fn get_entry(&self, id: u64) -> Option<&ObjectIndexEntry> {
        self.entries.get(&id)
    }

    /// Get an object entry through the validated identity API.
    pub fn entry(&self, object_id: ObjectId) -> Option<&ObjectIndexEntry> {
        self.entries.get(&object_id.get())
    }

    /// Get all objects in a specific fragment
    pub fn get_fragment_objects(&self, fragment_name: &str) -> Option<&Vec<u64>> {
        self.fragment_objects.get(fragment_name)
    }

    /// Get all object IDs
    pub fn all_object_ids(&self) -> Vec<u64> {
        self.entries.keys().cloned().collect()
    }

    /// Get all indexed object identities in deterministic numeric order.
    pub fn object_ids(&self) -> Result<Vec<ObjectId>> {
        let mut object_ids: Vec<_> = self
            .entries
            .keys()
            .copied()
            .map(ObjectId::try_from)
            .collect::<std::result::Result<_, _>>()
            .map_err(|_| Error::Archive("object index contains a null object identifier".into()))?;
        object_ids.sort_unstable();
        Ok(object_ids)
    }

    /// Get typed object identities for one fragment in source order.
    pub fn fragment_object_ids(&self, fragment_name: &str) -> Result<Option<Vec<ObjectId>>> {
        self.get_fragment_objects(fragment_name)
            .map(|ids| {
                ids.iter()
                    .copied()
                    .map(ObjectId::try_from)
                    .collect::<std::result::Result<Vec<_>, _>>()
                    .map_err(|_| {
                        Error::Archive(format!(
                            "fragment {fragment_name} contains a null object identifier"
                        ))
                    })
            })
            .transpose()
    }

    /// Get all entries
    pub fn all_entries(&self) -> Vec<&ObjectIndexEntry> {
        self.entries.values().collect()
    }

    /// Find objects by type
    pub fn find_objects_by_type(&self, object_type: u32) -> Vec<&ObjectIndexEntry> {
        self.entries
            .values()
            .filter(|entry| entry.object_type == object_type)
            .collect()
    }

    /// Get the reference graph for advanced queries
    ///
    /// The reference graph contains bidirectional relationships between objects,
    /// enabling queries like:
    /// - What objects does this reference? (outgoing edges)
    /// - What objects reference this? (incoming edges)
    /// - Find all dependencies of an object
    /// - Detect circular references
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let index = ObjectIndex::from_bundle(&bundle)?;
    /// let graph = index.reference_graph();
    ///
    /// // Find what a table references
    /// if let Some(refs) = graph.get_outgoing_refs(table_id) {
    ///     println!("Table references {} objects", refs.len());
    /// }
    ///
    /// // Find what references a style
    /// if let Some(refs) = graph.get_incoming_refs(style_id) {
    ///     println!("{} objects use this style", refs.len());
    /// }
    /// ```
    pub fn reference_graph(&self) -> &ReferenceGraph {
        &self.reference_graph
    }

    /// Get objects that are referenced by the given object
    ///
    /// Returns the "dependencies" of an object - all objects it points to.
    ///
    /// # Arguments
    ///
    /// * `object_id` - The source object ID
    ///
    /// # Returns
    ///
    /// Optional slice of referenced object IDs, or None if object has no outgoing references
    pub fn get_dependencies(&self, object_id: u64) -> Option<&[u64]> {
        self.reference_graph
            .get_outgoing_refs(object_id)
            .map(|v| v.as_slice())
    }

    /// Get typed dependencies without exposing raw sentinel IDs.
    pub fn dependencies(&self, object_id: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.reference_graph.outgoing(object_id)
    }

    /// Get objects that reference the given object
    ///
    /// Returns the "dependents" of an object - all objects that point to it.
    ///
    /// # Arguments
    ///
    /// * `object_id` - The target object ID
    ///
    /// # Returns
    ///
    /// Optional slice of referencing object IDs, or None if no objects reference this one
    pub fn get_dependents(&self, object_id: u64) -> Option<&[u64]> {
        self.reference_graph
            .get_incoming_refs(object_id)
            .map(|v| v.as_slice())
    }

    /// Get typed dependents without exposing raw sentinel IDs.
    pub fn dependents(&self, object_id: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.reference_graph.incoming(object_id)
    }

    /// Check if there are any circular references starting from the given object
    ///
    /// Performs iterative depth-first search to detect cycles in the reference graph.
    /// This is useful for validating document integrity.
    ///
    /// # Arguments
    ///
    /// * `object_id` - The starting object ID
    ///
    /// # Returns
    ///
    /// true if a cycle is detected, false otherwise
    ///
    /// # Performance
    ///
    /// O(V + E) where V is vertices and E is edges in the reachable subgraph
    pub fn has_circular_reference(&self, object_id: u64) -> bool {
        self.reference_graph.has_cycle_from(object_id)
    }

    /// Check for a cycle through the validated identity API.
    pub fn has_cycle_from(&self, object_id: ObjectId) -> bool {
        self.reference_graph.snapshot().has_cycle_from(object_id)
    }

    /// Get all objects reachable from the given object
    ///
    /// Performs breadth-first traversal to find all transitively referenced objects.
    /// Useful for extracting complete sub-documents or determining what needs
    /// to be loaded to fully resolve an object.
    ///
    /// # Arguments
    ///
    /// * `object_id` - The starting object ID
    ///
    /// # Returns
    ///
    /// Vector of all reachable object IDs (including the start object)
    ///
    /// # Performance
    ///
    /// O(V + E) where V is vertices and E is edges in the reachable subgraph
    pub fn get_transitive_dependencies(&self, object_id: u64) -> Vec<u64> {
        self.reference_graph.get_reachable(object_id)
    }

    /// Get typed transitive dependencies, including the starting object.
    pub fn reachable_from(&self, object_id: ObjectId) -> Vec<ObjectId> {
        self.reference_graph.snapshot().reachable(object_id)
    }

    /// Resolve an object reference to get the actual object data
    ///
    /// This is a key function for navigating the iWork document object graph.
    /// Objects reference each other by ID, and this function resolves those
    /// references to get the actual object data.
    ///
    /// # Arguments
    ///
    /// * `bundle` - The document bundle containing all archives
    /// * `object_id` - The ID of the object to resolve
    ///
    /// # Returns
    ///
    /// * `Ok(Some(ResolvedObject))` - The resolved object with all its data
    /// * `Ok(None)` - Object ID not found in index
    /// * `Err(_)` - Archive file not found or other error
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// // Resolve a table's data_store reference
    /// if let Some(data_store) = index.resolve_object(&bundle, data_store_id)? {
    ///     // Parse the TableDataList to get cell values
    ///     for msg in &data_store.messages {
    ///         // Process message data
    ///     }
    /// }
    /// ```
    pub fn resolve_object(
        &self,
        bundle: &Bundle,
        object_id: u64,
    ) -> Result<Option<ResolvedObject>> {
        let Some(entry) = self.get_entry(object_id) else {
            return Ok(None);
        };

        let Some(archive) = bundle.get_archive(&entry.fragment_name) else {
            return Err(Error::Bundle(format!(
                "Archive {} not found",
                entry.fragment_name
            )));
        };

        // Find the object in the archive
        for object in &archive.objects {
            if object.archive_info.identifier == Some(object_id) {
                return Ok(Some(ResolvedObject {
                    id: object_id,
                    archive_info: object.archive_info.clone(),
                    messages: object.messages.clone(),
                }));
            }
        }

        Ok(None)
    }

    /// Resolve an object through the validated identity API.
    pub fn resolve(&self, bundle: &Bundle, object_id: ObjectId) -> Result<Option<ResolvedObject>> {
        self.resolve_object(bundle, object_id.get())
    }

    /// Batch resolve multiple object references
    ///
    /// More efficient than calling `resolve_object` multiple times
    /// as it minimizes archive lookups.
    ///
    /// # Arguments
    ///
    /// * `bundle` - The document bundle
    /// * `object_ids` - Slice of object IDs to resolve
    ///
    /// # Returns
    ///
    /// Vector of successfully resolved objects (may be smaller than input if some IDs don't exist)
    pub fn resolve_objects(
        &self,
        bundle: &Bundle,
        object_ids: &[u64],
    ) -> Result<Vec<ResolvedObject>> {
        let mut resolved = Vec::with_capacity(object_ids.len());

        // Group object IDs by their archive to minimize archive lookups
        let mut objects_by_archive: std::collections::HashMap<&str, Vec<u64>> =
            std::collections::HashMap::new();

        for &object_id in object_ids {
            if let Some(entry) = self.get_entry(object_id) {
                objects_by_archive
                    .entry(&entry.fragment_name)
                    .or_default()
                    .push(object_id);
            }
        }

        // Resolve objects archive by archive
        for (archive_name, ids) in objects_by_archive {
            if let Some(archive) = bundle.get_archive(archive_name) {
                for object in &archive.objects {
                    if let Some(obj_id) = object.archive_info.identifier
                        && ids.contains(&obj_id)
                    {
                        resolved.push(ResolvedObject {
                            id: obj_id,
                            archive_info: object.archive_info.clone(),
                            messages: object.messages.clone(),
                        });
                    }
                }
            }
        }

        Ok(resolved)
    }

    /// Batch-resolve objects through the validated identity API.
    pub fn resolve_many(
        &self,
        bundle: &Bundle,
        object_ids: &[ObjectId],
    ) -> Result<Vec<ResolvedObject>> {
        let raw_ids: Vec<_> = object_ids.iter().map(|object_id| object_id.get()).collect();
        let mut resolved = self.resolve_objects(bundle, &raw_ids)?;
        let order: HashMap<_, _> = object_ids
            .iter()
            .enumerate()
            .map(|(position, object_id)| (object_id.get(), position))
            .collect();
        resolved
            .sort_unstable_by_key(|object| order.get(&object.id).copied().unwrap_or(usize::MAX));
        Ok(resolved)
    }

    /// Resolve an object and its typed dependency closure.
    pub fn resolve_reachable(
        &self,
        bundle: &Bundle,
        object_id: ObjectId,
    ) -> Result<Vec<ResolvedObject>> {
        let object_ids = self.reachable_from(object_id);
        self.resolve_many(bundle, &object_ids)
    }

    /// Resolve an object and all its dependencies transitively
    ///
    /// This performs a breadth-first traversal of the object graph,
    /// resolving the given object and all objects it references.
    ///
    /// # Arguments
    ///
    /// * `bundle` - The document bundle
    /// * `object_id` - The root object ID to start resolving from
    ///
    /// # Returns
    ///
    /// Vector of all resolved objects reachable from the root
    ///
    /// # Performance
    ///
    /// O(V + E) where V is the number of reachable objects and E is edges.
    /// Uses batch resolution to minimize archive lookups.
    pub fn resolve_with_dependencies(
        &self,
        bundle: &Bundle,
        object_id: u64,
    ) -> Result<Vec<ResolvedObject>> {
        let all_ids = self.get_transitive_dependencies(object_id);
        self.resolve_objects(bundle, &all_ids)
    }

    /// Check if an object exists in the index
    pub fn contains_object(&self, object_id: u64) -> bool {
        self.entries.contains_key(&object_id)
    }

    /// Check for an indexed object through the validated identity API.
    pub fn contains(&self, object_id: ObjectId) -> bool {
        self.entries.contains_key(&object_id.get())
    }

    /// Get the total number of indexed objects
    pub fn object_count(&self) -> usize {
        self.entries.len()
    }

    /// Get the number of fragments (IWA files) in the index
    pub fn fragment_count(&self) -> usize {
        self.fragment_objects.len()
    }

    /// Get statistics about the object index
    pub fn stats(&self) -> ObjectIndexStats {
        let total_objects = self.entries.len();
        let total_fragments = self.fragment_objects.len();
        let total_references = self.reference_graph.edge_count();
        let avg_refs_per_object = if total_objects > 0 {
            total_references as f64 / total_objects as f64
        } else {
            0.0
        };

        ObjectIndexStats {
            total_objects,
            total_fragments,
            total_references,
            avg_refs_per_object,
        }
    }

    #[allow(deprecated)]
    fn extract_legacy_sheet_headers(
        &mut self,
        object_id: u64,
        sheet: &crate::protobuf::tn::SheetArchive,
    ) {
        if let Some(header) = &sheet.header_storage {
            self.extract_reference(object_id, header);
        }
        if let Some(footer) = &sheet.footer_storage {
            self.extract_reference(object_id, footer);
        }
    }
}

fn extract_table_data_list_entry_references(
    index: &mut ObjectIndex,
    object_id: u64,
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
        index.extract_reference(object_id, reference);
    }
}

/// Statistics about the object index
#[derive(Debug, Clone)]
pub struct ObjectIndexStats {
    /// Total number of objects in the index
    pub total_objects: usize,
    /// Total number of IWA fragments
    pub total_fragments: usize,
    /// Total number of object references
    pub total_references: usize,
    /// Average references per object
    pub avg_refs_per_object: f64,
}

/// A resolved object with its full data
#[derive(Debug, Clone)]
pub struct ResolvedObject {
    /// Object identifier
    pub id: u64,
    /// Archive information
    pub archive_info: crate::archive::ArchiveInfo,
    /// Raw message data
    pub messages: Vec<RawMessage>,
}

impl ResolvedObject {
    /// Return the validated object identity, if the compatibility payload is
    /// non-null.
    pub fn object_id(&self) -> Option<ObjectId> {
        ObjectId::new(self.id)
    }

    /// Get the primary message type
    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|msg| msg.type_)
    }

    /// Get all message types
    pub fn message_types(&self) -> Vec<u32> {
        self.messages.iter().map(|msg| msg.type_).collect()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject, RawMessage};
    use crate::protobuf::tp::{DocumentArchive, SectionArchive, SectionTemplateArchive};
    use crate::protobuf::tsp::Reference;
    use crate::protobuf::tst::{self, TableDataList, TableDataListSegment};
    use prost::Message;

    #[test]
    fn test_object_index_creation() {
        let index = ObjectIndex::new();
        assert!(index.entries.is_empty());
        assert!(index.fragment_objects.is_empty());
    }

    #[test]
    fn test_object_index_entry() {
        let entry = ObjectIndexEntry {
            id: 123,
            fragment_name: "Document.iwa".to_string(),
            data_offset: 100,
            data_length: 200,
            object_type: 42,
        };

        assert_eq!(entry.id, 123);
        assert_eq!(entry.fragment_name, "Document.iwa");
        assert_eq!(entry.object_type, 42);
        assert_eq!(entry.object_id(), ObjectId::new(123));
    }

    #[test]
    fn test_object_index_with_reference_graph() {
        let index = ObjectIndex::new();

        assert!(index.reference_graph().is_empty());
        assert_eq!(index.get_dependencies(1), None);
        assert_eq!(index.get_dependents(1), None);
        assert!(!index.has_circular_reference(1));
        assert_eq!(index.get_transitive_dependencies(1), vec![1]);

        let object_id = ObjectId::try_from(1).unwrap();
        assert!(index.dependencies(object_id).is_none());
        assert!(index.dependents(object_id).is_none());
        assert!(!index.has_cycle_from(object_id));
        assert_eq!(index.reachable_from(object_id), vec![object_id]);
    }

    #[test]
    fn indexes_authoritative_message_info_references() {
        let mut object = ArchiveObject::new(
            10,
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        object.archive_info.message_infos[0].object_references = vec![0, 20, 30, 20, 0];
        let archive = Archive {
            objects: vec![object],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Test.iwa", &archive).unwrap();

        assert_eq!(index.get_dependencies(10), Some([20, 30].as_slice()));
        assert_eq!(index.get_dependents(20), Some([10].as_slice()));
        assert_eq!(index.stats().total_references, 2);
    }

    #[test]
    fn authoritative_null_only_references_suppress_legacy_fallback() {
        let table_data = TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            entries: Vec::new(),
            segments: vec![Reference {
                identifier: 20,
                ..Default::default()
            }],
            ..Default::default()
        };
        let mut object = ArchiveObject::new(
            10,
            vec![RawMessage {
                type_: 6005,
                data: table_data.encode_to_vec(),
            }],
        )
        .unwrap();
        object.archive_info.message_infos[0].object_references = vec![0];

        let archive = Archive {
            objects: vec![object],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Test.iwa", &archive).unwrap();

        assert_eq!(index.get_dependencies(10), None);
        assert_eq!(index.stats().total_references, 0);
    }

    #[test]
    fn typed_object_index_queries_preserve_order_and_identity() {
        let mut object = ArchiveObject::new(
            10,
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        object.archive_info.message_infos[0].object_references = vec![20, 30, 20];
        let archive = Archive {
            objects: vec![object],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Test.iwa", &archive).unwrap();

        let source = ObjectId::try_from(10).unwrap();
        let target = ObjectId::try_from(20).unwrap();

        assert_eq!(
            index.entry(source).and_then(ObjectIndexEntry::object_id),
            Some(source)
        );
        assert_eq!(index.object_ids().unwrap(), vec![source]);
        assert_eq!(
            index.fragment_object_ids("Index/Test.iwa").unwrap(),
            Some(vec![source])
        );
        assert_eq!(index.fragment_object_ids("missing.iwa").unwrap(), None);
        assert_eq!(
            index.dependencies(source).unwrap().collect::<Vec<_>>(),
            vec![target, ObjectId::try_from(30).unwrap()]
        );
        assert_eq!(
            index.dependents(target).unwrap().collect::<Vec<_>>(),
            vec![source]
        );
        assert_eq!(
            index.reachable_from(source),
            vec![source, target, ObjectId::try_from(30).unwrap()]
        );
        assert!(!index.has_cycle_from(source));
        assert!(index.contains(source));
    }

    #[test]
    fn rejects_null_archive_object_ids() {
        let object = ArchiveObject::new(
            0,
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let archive = Archive {
            objects: vec![object],
        };

        let error = ObjectIndex::new()
            .parse_archive("Index/Test.iwa", &archive)
            .unwrap_err();
        assert!(
            matches!(error, Error::Archive(message) if message.contains("null object identifier"))
        );
    }

    #[test]
    fn rejects_missing_archive_object_ids() {
        let mut object = ArchiveObject::new(
            10,
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        object.archive_info.identifier = None;
        let archive = Archive {
            objects: vec![object],
        };

        let error = ObjectIndex::new()
            .parse_archive("Index/Test.iwa", &archive)
            .unwrap_err();
        assert!(
            matches!(error, Error::Archive(message) if message.contains("without an identifier"))
        );
    }

    #[test]
    fn fallback_indexes_segmented_table_data_list_references() {
        let root = TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            next_list_id: 2,
            entries: Vec::new(),
            segments: vec![Reference {
                identifier: 20,
                ..Default::default()
            }],
            is_new_for_bnc: Some(true),
        };
        let segment = TableDataListSegment {
            list_type: root.list_type,
            key_range: crate::protobuf::tsp::Range {
                location: 1,
                length: 1,
            },
            entries: vec![tst::table_data_list::ListEntry {
                key: 1,
                refcount: 1,
                rich_text_payload: Some(Reference {
                    identifier: 30,
                    ..Default::default()
                }),
                ..Default::default()
            }],
        };
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    10,
                    vec![RawMessage {
                        type_: 6005,
                        data: root.encode_to_vec(),
                    }],
                )
                .unwrap(),
                ArchiveObject::new(
                    20,
                    vec![RawMessage {
                        type_: 6011,
                        data: segment.encode_to_vec(),
                    }],
                )
                .unwrap(),
            ],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Test.iwa", &archive).unwrap();
        assert_eq!(index.get_dependencies(10), Some([20].as_slice()));
        assert_eq!(index.get_dependencies(20), Some([30].as_slice()));
    }

    #[test]
    fn fallback_indexes_comment_author_and_replies() {
        let comment = crate::protobuf::tsd::CommentStorageArchive {
            author: Some(Reference {
                identifier: 20,
                ..Default::default()
            }),
            replies: vec![Reference {
                identifier: 30,
                ..Default::default()
            }],
            ..Default::default()
        };
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    10,
                    vec![RawMessage {
                        type_: 3056,
                        data: comment.encode_to_vec(),
                    }],
                )
                .unwrap(),
            ],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Comments.iwa", &archive).unwrap();
        assert_eq!(index.get_dependencies(10), Some([20, 30].as_slice()));
    }

    #[test]
    fn pages_fallback_indexes_document_section_and_template_graph() {
        let reference = |identifier| Reference {
            identifier,
            ..Default::default()
        };
        let document = DocumentArchive {
            body_storage: Some(reference(42)),
            section: Some(reference(43)),
            theme: Some(reference(44)),
            page_templates: vec![reference(45)],
            ..Default::default()
        };
        let section = SectionArchive {
            first_section_template_page: Some(reference(50)),
            even_section_template_page: Some(reference(51)),
            odd_section_template_page: Some(reference(52)),
            user_defined_guide_storage: Some(reference(53)),
            ..Default::default()
        };
        let template = SectionTemplateArchive {
            headers: vec![reference(60)],
            footers: vec![reference(61)],
            section_template_drawables: vec![reference(62)],
            ..Default::default()
        };
        let object = |identifier, type_, data| {
            ArchiveObject::new(identifier, vec![RawMessage { type_, data }]).unwrap()
        };
        let archive = Archive {
            objects: vec![
                object(1, 10000, document.encode_to_vec()),
                object(43, 10011, section.encode_to_vec()),
                object(50, 10143, template.encode_to_vec()),
            ],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Document.iwa", &archive).unwrap();

        let document_dependencies = index.get_dependencies(1).unwrap();
        for identifier in [42, 43, 44, 45] {
            assert!(document_dependencies.contains(&identifier));
        }
        let section_dependencies = index.get_dependencies(43).unwrap();
        for identifier in [50, 51, 52, 53] {
            assert!(section_dependencies.contains(&identifier));
        }
        let template_dependencies = index.get_dependencies(50).unwrap();
        assert_eq!(template_dependencies, [60, 61, 62]);
    }
}
