//! Object Index for Cross-Referencing in iWork Documents
//!
//! iWork documents contain an object index that maps object IDs to their
//! locations in IWA files. This allows objects to reference each other
//! across different archive files.

use std::collections::HashMap;

use crate::iwa::archive::{Archive, ArchiveObject, RawMessage};
use crate::iwa::bundle::Bundle;
use crate::iwa::ref_graph::ReferenceGraph;
use crate::iwa::{Error, Result};

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

        // Look for the object index, typically in Metadata.iwa or a similar file
        if let Some(metadata_archive) = bundle.get_archive("Index/Metadata.iwa") {
            index.parse_metadata_archive(metadata_archive)?;
        }

        // Parse all archives to build the index
        for (archive_name, archive) in bundle.archives() {
            index.parse_archive(archive_name, archive)?;
        }

        Ok(index)
    }

    /// Parse the metadata archive to find object references
    fn parse_metadata_archive(&mut self, archive: &Archive) -> Result<()> {
        for object in &archive.objects {
            if let Some(identifier) = object.archive_info.identifier {
                // Look for object references in message data
                self.parse_object_references(identifier, object)?;
            }
        }
        Ok(())
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
            if let Some(identifier) = object.archive_info.identifier {
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
                        crate::iwa::protobuf::tst::TableModelArchive::decode(&*raw_msg.data)
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
                        self.extract_reference(object_id, &table.data_store.column_headers);
                        self.extract_reference(object_id, &table.data_store.string_table);
                        self.extract_reference(object_id, &table.data_store.style_table);
                        self.extract_reference(object_id, &table.data_store.formula_table);
                        self.extract_reference(object_id, &table.data_store.format_table);

                        // Optional references
                        if let Some(ref formula_error_table) = table.data_store.formula_error_table
                        {
                            self.extract_reference(object_id, formula_error_table);
                        }
                        if let Some(ref choice_list) =
                            table.data_store.multiple_choice_list_format_table
                        {
                            self.extract_reference(object_id, choice_list);
                        }
                        if let Some(ref merge_map) = table.data_store.merge_region_map {
                            self.extract_reference(object_id, merge_map);
                        }
                    }
                },

                6005 | 6201 => {
                    // TST.TableDataList - may contain references to other data structures
                    // The actual cell data is stored here
                },

                // TSWP (Word Processing/Text) types
                2001..=2022 => {
                    // TSWP.StorageArchive contains text content and may reference styles
                    if let Ok(storage) =
                        crate::iwa::protobuf::tswp::StorageArchive::decode(&*raw_msg.data)
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
                    if let Ok(slide) =
                        crate::iwa::protobuf::kn::SlideArchive::decode(&*raw_msg.data)
                    {
                        // Extract style reference
                        self.extract_reference(object_id, &slide.style);

                        // Extract drawable references (shapes, images, text boxes)
                        for drawable in &slide.drawables {
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
                    if let Ok(show) = crate::iwa::protobuf::kn::ShowArchive::decode(&*raw_msg.data)
                    {
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
                    if let Ok(sheet) =
                        crate::iwa::protobuf::tn::SheetArchive::decode(&*raw_msg.data)
                    {
                        // Extract drawable info references
                        for drawable_ref in &sheet.drawable_infos {
                            self.extract_reference(object_id, drawable_ref);
                        }

                        // Extract header/footer storage references if present
                        if let Some(ref header) = sheet.header_storage {
                            self.extract_reference(object_id, header);
                        }
                        if let Some(ref footer) = sheet.footer_storage {
                            self.extract_reference(object_id, footer);
                        }
                    }
                },

                // TSD (Drawing/Shape) types
                // Implementation Status: ✓ COMPLETED (2025-11-04)
                // Based on TSDArchives.proto and libetonyek's reference extraction
                3002 => {
                    // TSD.DrawableArchive - base type for all drawables
                    if let Ok(drawable) =
                        crate::iwa::protobuf::tsd::DrawableArchive::decode(&*raw_msg.data)
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
                        crate::iwa::protobuf::tsd::ContainerArchive::decode(&*raw_msg.data)
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
                    if let Ok(shape) =
                        crate::iwa::protobuf::tsd::ShapeArchive::decode(&*raw_msg.data)
                    {
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
                    if let Ok(image) =
                        crate::iwa::protobuf::tsd::ImageArchive::decode(&*raw_msg.data)
                    {
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
                    if let Ok(mask) = crate::iwa::protobuf::tsd::MaskArchive::decode(&*raw_msg.data)
                    {
                        // Extract parent from super DrawableArchive (required field)
                        if let Some(ref parent) = mask.super_.parent {
                            self.extract_reference(object_id, parent);
                        }
                        // Note: pathsource is embedded data, not a reference
                    }
                },
                3007 => {
                    // TSD.MovieArchive - video objects
                    if let Ok(movie) =
                        crate::iwa::protobuf::tsd::MovieArchive::decode(&*raw_msg.data)
                    {
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
                    if let Ok(group) =
                        crate::iwa::protobuf::tsd::GroupArchive::decode(&*raw_msg.data)
                    {
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
                        crate::iwa::protobuf::tsd::ConnectionLineArchive::decode(&*raw_msg.data)
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

                // TSCH (Chart) types
                // Implementation Status: ✓ COMPLETED (2025-11-04)
                // Based on TSCHArchives.proto and libetonyek's chart parsing
                5000 => {
                    // TSCH.PreUFF.ChartInfoArchive - legacy chart format
                    // This is a pre-unified format chart, structure may vary
                    // Attempt basic reference extraction but may fail gracefully
                    if let Ok(chart_info) =
                        crate::iwa::protobuf::tsch::pre_uff::ChartInfoArchive::decode(
                            &*raw_msg.data,
                        )
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
                        crate::iwa::protobuf::tsch::ChartMediatorArchive::decode(&*raw_msg.data)
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
                        crate::iwa::protobuf::tsch::ChartStylePreset::decode(&*raw_msg.data)
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
                        crate::iwa::protobuf::tsch::ChartDrawableArchive::decode(&*raw_msg.data)
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
                    if let Ok(doc) =
                        crate::iwa::protobuf::tp::DocumentArchive::decode(&*raw_msg.data)
                    {
                        // Extract theme reference
                        if let Some(ref theme) = doc.theme {
                            self.extract_reference(object_id, theme);
                        }

                        // Extract stylesheet reference
                        if let Some(ref stylesheet) = doc.stylesheet {
                            self.extract_reference(object_id, stylesheet);
                        }
                    }
                },

                10011 => {
                    // TP.SectionArchive
                    // Note: SectionArchive has a complex structure
                    // References are embedded in nested structures
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
    fn extract_reference(
        &mut self,
        source_id: u64,
        reference: &crate::iwa::protobuf::tsp::Reference,
    ) {
        let target_id = reference.identifier;

        // Ignore null/zero references (0 typically means "no reference")
        if target_id == 0 {
            return;
        }

        // Build the reference graph: track both outgoing and incoming references
        // This enables bidirectional graph traversal
        self.reference_graph.add_reference(source_id, target_id);
    }

    /// Get an object entry by ID
    pub fn get_entry(&self, id: u64) -> Option<&ObjectIndexEntry> {
        self.entries.get(&id)
    }

    /// Get all objects in a specific fragment
    pub fn get_fragment_objects(&self, fragment_name: &str) -> Option<&Vec<u64>> {
        self.fragment_objects.get(fragment_name)
    }

    /// Get all object IDs
    pub fn all_object_ids(&self) -> Vec<u64> {
        self.entries.keys().cloned().collect()
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

    /// Check if there are any circular references starting from the given object
    ///
    /// Performs depth-first search to detect cycles in the reference graph.
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
    pub archive_info: crate::iwa::archive::ArchiveInfo,
    /// Raw message data
    pub messages: Vec<RawMessage>,
}

impl ResolvedObject {
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
    }

    #[test]
    fn test_object_index_with_reference_graph() {
        let index = ObjectIndex::new();

        assert!(index.reference_graph().is_empty());
        assert_eq!(index.get_dependencies(1), None);
        assert_eq!(index.get_dependents(1), None);
        assert!(!index.has_circular_reference(1));
        assert_eq!(index.get_transitive_dependencies(1), vec![1]);
    }
}
