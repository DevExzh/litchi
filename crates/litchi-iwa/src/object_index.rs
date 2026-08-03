//! Object Index for Cross-Referencing in iWork Documents
//!
//! iWork documents contain an object index that maps object IDs to their
//! locations in IWA files. This allows objects to reference each other
//! across different archive files.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use crate::archive::{Archive, RawMessage};
use crate::bundle::Bundle;
use crate::ref_graph::{ObjectId, ObjectIdIter, ReferenceGraph};
use crate::{Error, Result};

mod reference_extraction;

/// Represents an entry in the object index
#[derive(Debug, Clone)]
pub struct ObjectIndexEntry {
    /// Unique, validated object identifier.
    id: ObjectId,
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
    /// Return the validated object identity.
    pub const fn id(&self) -> ObjectId {
        self.id
    }

    /// Return the validated object identity, if this compatibility entry is
    /// non-null.
    pub fn object_id(&self) -> Option<ObjectId> {
        Some(self.id)
    }
}

/// Object index that maps object IDs to their locations
#[derive(Debug, Clone)]
pub struct ObjectIndex {
    /// Map from object ID to index entry
    entries: Arc<HashMap<u64, ObjectIndexEntry>>,
    /// Map from fragment name to list of object IDs
    fragment_objects: Arc<HashMap<String, Vec<u64>>>,
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
            entries: Arc::new(HashMap::new()),
            fragment_objects: Arc::new(HashMap::new()),
            reference_graph: ReferenceGraph::new(),
        }
    }

    /// Build object index from a bundle
    pub fn from_bundle(bundle: &Bundle) -> Result<Self> {
        let mut index = Self::new();

        // Bundle traversal is already sorted at ingress, so fragment and
        // reverse-reference order do not depend on randomized map storage.
        for (archive_name, archive) in bundle.iter_archives() {
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

            if let Some(existing) = self.entries.get(&identifier) {
                return Err(Error::Archive(format!(
                    "object {identifier} occurs in archives {} and {archive_name}",
                    existing.fragment_name
                )));
            }

            // Determine object type from first message
            let object_type = object.messages.first().map(|msg| msg.type_).unwrap_or(0);

            let entry = ObjectIndexEntry {
                id: object_id,
                fragment_name: archive_name.to_string(),
                // Use actual byte offsets from the parsed archive
                // These match the approach used in libetonyek's ObjectRecord
                data_offset: object.data_offset,
                data_length: object.data_length,
                object_type,
            };

            Arc::make_mut(&mut self.entries).insert(identifier, entry);
            Arc::make_mut(&mut self.fragment_objects)
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
                reference_extraction::extract(object_id, object, &mut self.reference_graph)?;
            }
        }
        Ok(())
    }

    /// Get an object entry by ID
    #[deprecated(note = "use entry(ObjectId) for checked identity semantics")]
    pub fn get_entry(&self, id: u64) -> Option<&ObjectIndexEntry> {
        self.entries.get(&id)
    }

    /// Get an object entry through the validated identity API.
    pub fn entry(&self, object_id: ObjectId) -> Option<&ObjectIndexEntry> {
        self.entries.get(&object_id.get())
    }

    /// Get all objects in a specific fragment
    #[deprecated(note = "use fragment_object_ids for checked identity semantics")]
    pub fn get_fragment_objects(&self, fragment_name: &str) -> Option<&Vec<u64>> {
        self.fragment_objects.get(fragment_name)
    }

    /// Get all object IDs in deterministic numeric order.
    #[deprecated(note = "use object_ids for checked identity semantics")]
    pub fn all_object_ids(&self) -> Vec<u64> {
        let mut object_ids: Vec<_> = self.entries.keys().copied().collect();
        object_ids.sort_unstable();
        object_ids
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
        self.fragment_objects
            .get(fragment_name)
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

    /// Get all entries in deterministic numeric object-ID order.
    pub fn all_entries(&self) -> Vec<&ObjectIndexEntry> {
        let mut entries: Vec<_> = self.entries.values().collect();
        entries.sort_unstable_by_key(|entry| entry.id());
        entries
    }

    /// Find objects by type in deterministic numeric object-ID order.
    pub fn find_objects_by_type(&self, object_type: u32) -> Vec<&ObjectIndexEntry> {
        let mut entries: Vec<_> = self
            .entries
            .values()
            .filter(|entry| entry.object_type == object_type)
            .collect();
        entries.sort_unstable_by_key(|entry| entry.id());
        entries
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
    /// if let Some(refs) = graph.outgoing(table_id) {
    ///     println!("Table references {} objects", refs.len());
    /// }
    ///
    /// // Find what references a style
    /// if let Some(refs) = graph.incoming(style_id) {
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
    /// Optional owned compatibility view of referenced object IDs.
    ///
    /// Prefer [`Self::dependencies`] for the allocation-free typed view.
    #[deprecated(note = "use dependencies(ObjectId) for the typed, allocation-free view")]
    pub fn get_dependencies(&self, object_id: u64) -> Option<Vec<u64>> {
        let object_id = ObjectId::new(object_id)?;
        self.dependencies(object_id)
            .map(|references| references.map(ObjectId::get).collect())
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
    /// Optional owned compatibility view of referencing object IDs.
    ///
    /// Prefer [`Self::dependents`] for the allocation-free typed view.
    #[deprecated(note = "use dependents(ObjectId) for the typed, allocation-free view")]
    pub fn get_dependents(&self, object_id: u64) -> Option<Vec<u64>> {
        let object_id = ObjectId::new(object_id)?;
        self.dependents(object_id)
            .map(|references| references.map(ObjectId::get).collect())
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
    #[deprecated(note = "use has_cycle_from(ObjectId) for checked identity semantics")]
    pub fn has_circular_reference(&self, object_id: u64) -> bool {
        ObjectId::new(object_id).is_some_and(|object_id| self.has_cycle_from(object_id))
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
    #[deprecated(note = "use reachable_from(ObjectId) for the typed traversal")]
    pub fn get_transitive_dependencies(&self, object_id: u64) -> Vec<u64> {
        let Some(object_id) = ObjectId::new(object_id) else {
            return Vec::new();
        };
        self.reachable_from(object_id)
            .into_iter()
            .map(ObjectId::get)
            .collect()
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
    /// if let Some(data_store) = index.resolve(&bundle, data_store_id)? {
    ///     // Parse the TableDataList to get cell values
    ///     for msg in &data_store.messages {
    ///         // Process message data
    ///     }
    /// }
    /// ```
    #[deprecated(note = "use resolve(ObjectId) for checked identity semantics")]
    pub fn resolve_object(
        &self,
        bundle: &Bundle,
        object_id: u64,
    ) -> Result<Option<ResolvedObject>> {
        self.resolve_id(bundle, object_id)
    }

    /// Resolve a protobuf wire identifier for crate-internal readers.
    ///
    /// The raw numeric form is kept behind the crate boundary because parsed
    /// protobuf references enter this module as `u64`. Public callers should
    /// validate once and use [`Self::resolve`] with [`ObjectId`].
    pub(crate) fn resolve_id(
        &self,
        bundle: &Bundle,
        object_id: u64,
    ) -> Result<Option<ResolvedObject>> {
        let Some(object_id) = ObjectId::new(object_id) else {
            return Ok(None);
        };
        self.resolve(bundle, object_id)
    }

    /// Resolve an object through the validated identity API.
    pub fn resolve(&self, bundle: &Bundle, object_id: ObjectId) -> Result<Option<ResolvedObject>> {
        let Some(entry) = self.entry(object_id) else {
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
            if object.archive_info.identifier == Some(object_id.get()) {
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
    /// More efficient than calling `resolve` multiple times
    /// as it minimizes archive lookups.
    ///
    /// # Arguments
    ///
    /// * `bundle` - The document bundle
    /// * `object_ids` - Slice of object IDs to resolve
    ///
    /// # Returns
    ///
    /// Vector of successfully resolved objects in the caller's input order.
    /// The result may be smaller than the input if some IDs do not exist.
    #[deprecated(note = "use resolve_many(&[ObjectId]) for checked identity semantics")]
    pub fn resolve_objects(
        &self,
        bundle: &Bundle,
        object_ids: &[u64],
    ) -> Result<Vec<ResolvedObject>> {
        let object_ids = object_ids
            .iter()
            .copied()
            .filter_map(ObjectId::new)
            .collect::<Vec<_>>();
        self.resolve_many_inner(bundle, &object_ids)
    }

    fn resolve_many_inner(
        &self,
        bundle: &Bundle,
        object_ids: &[ObjectId],
    ) -> Result<Vec<ResolvedObject>> {
        // Group object IDs by their archive to minimize archive lookups
        let mut objects_by_archive: std::collections::HashMap<&str, HashSet<ObjectId>> =
            std::collections::HashMap::new();

        for &object_id in object_ids {
            if let Some(entry) = self.entry(object_id) {
                objects_by_archive
                    .entry(&entry.fragment_name)
                    .or_default()
                    .insert(object_id);
            }
        }

        let mut resolved_by_id = HashMap::with_capacity(object_ids.len());

        // Resolve objects archive by archive
        for (archive_name, ids) in objects_by_archive {
            if let Some(archive) = bundle.get_archive(archive_name) {
                for object in &archive.objects {
                    if let Some(obj_id) = object.archive_info.identifier
                        && let Ok(object_id) = ObjectId::try_from(obj_id)
                        && ids.contains(&object_id)
                    {
                        let resolved = ResolvedObject {
                            id: object_id,
                            archive_info: object.archive_info.clone(),
                            messages: object.messages.clone(),
                        };
                        if resolved_by_id.insert(object_id, resolved).is_some() {
                            return Err(Error::Archive(format!(
                                "object {obj_id} occurs in more than one archive"
                            )));
                        }
                    }
                }
            }
        }

        Ok(object_ids
            .iter()
            .filter_map(|object_id| resolved_by_id.remove(object_id))
            .collect())
    }

    /// Batch-resolve objects through the validated identity API.
    pub fn resolve_many(
        &self,
        bundle: &Bundle,
        object_ids: &[ObjectId],
    ) -> Result<Vec<ResolvedObject>> {
        let mut requested = HashSet::with_capacity(object_ids.len());
        for object_id in object_ids {
            if !requested.insert(*object_id) {
                return Err(Error::Archive(format!(
                    "object {object_id:?} occurs more than once in a batch"
                )));
            }
            if !self.contains(*object_id) {
                return Err(Error::Archive(format!(
                    "object {} is not present in the object index",
                    object_id.get()
                )));
            }
        }

        let resolved = self.resolve_many_inner(bundle, object_ids)?;
        let resolved_ids: HashSet<_> = resolved.iter().map(ResolvedObject::id).collect();
        if let Some(missing) = object_ids
            .iter()
            .find(|object_id| !resolved_ids.contains(object_id))
        {
            return Err(Error::Bundle(format!(
                "object {} could not be resolved from the bundle",
                missing.get()
            )));
        }
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
    #[deprecated(note = "use resolve_reachable(ObjectId) for checked identity semantics")]
    pub fn resolve_with_dependencies(
        &self,
        bundle: &Bundle,
        object_id: u64,
    ) -> Result<Vec<ResolvedObject>> {
        let Some(object_id) = ObjectId::new(object_id) else {
            return Ok(Vec::new());
        };
        let all_ids = self
            .reachable_from(object_id)
            .into_iter()
            .filter(|object_id| self.contains(*object_id))
            .collect::<Vec<_>>();
        self.resolve_many_inner(bundle, &all_ids)
    }

    /// Check if an object exists in the index
    #[deprecated(note = "use contains(ObjectId) for checked identity semantics")]
    pub fn contains_object(&self, object_id: u64) -> bool {
        ObjectId::new(object_id).is_some_and(|object_id| self.contains(object_id))
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
    /// Validated object identifier.
    id: ObjectId,
    /// Archive information
    pub archive_info: crate::archive::ArchiveInfo,
    /// Raw message data
    pub messages: Vec<RawMessage>,
}

impl ResolvedObject {
    /// Return the validated object identity.
    pub const fn id(&self) -> ObjectId {
        self.id
    }

    /// Return the validated object identity, if the compatibility payload is
    /// non-null.
    pub fn object_id(&self) -> Option<ObjectId> {
        Some(self.id)
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
#[allow(deprecated)]
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
    fn object_index_clones_share_maps_until_indexing_mutates() {
        let object = ArchiveObject::new(
            10,
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let mut index = ObjectIndex::new();
        index
            .parse_archive(
                "Index/First.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        let snapshot = index.clone();
        assert!(Arc::ptr_eq(&index.entries, &snapshot.entries));
        assert!(Arc::ptr_eq(
            &index.fragment_objects,
            &snapshot.fragment_objects
        ));

        let second = ArchiveObject::new(
            20,
            vec![RawMessage {
                type_: 43,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let mut edited = snapshot.clone();
        edited
            .parse_archive(
                "Index/Second.iwa",
                &Archive {
                    objects: vec![second],
                },
            )
            .unwrap();

        assert_eq!(index.object_count(), 1);
        assert_eq!(edited.object_count(), 2);
        assert!(!Arc::ptr_eq(&index.entries, &edited.entries));
        assert!(!Arc::ptr_eq(
            &index.fragment_objects,
            &edited.fragment_objects
        ));
    }

    #[test]
    fn object_indexes_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<ObjectIndex>();
    }

    #[test]
    fn test_object_index_entry() {
        let entry = ObjectIndexEntry {
            id: ObjectId::try_from(123).unwrap(),
            fragment_name: "Document.iwa".to_string(),
            data_offset: 100,
            data_length: 200,
            object_type: 42,
        };

        assert_eq!(entry.id().get(), 123);
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

        assert_eq!(index.get_dependencies(10), Some(vec![20, 30]));
        assert_eq!(index.get_dependents(20), Some(vec![10]));
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
    #[allow(deprecated)]
    fn compatibility_object_queries_are_deterministically_ordered() {
        let objects = [(30, 7), (10, 7), (20, 8)]
            .into_iter()
            .map(|(id, object_type)| {
                let mut object = ArchiveObject::new(
                    id,
                    vec![RawMessage {
                        type_: object_type,
                        data: Vec::new(),
                    }],
                )
                .unwrap();
                object.archive_info.message_infos[0].type_ = object_type;
                object
            })
            .collect();
        let mut index = ObjectIndex::new();
        index
            .parse_archive("Index/Test.iwa", &Archive { objects })
            .unwrap();

        assert_eq!(index.all_object_ids(), vec![10, 20, 30]);
        assert_eq!(
            index
                .all_entries()
                .into_iter()
                .map(|entry| entry.id().get())
                .collect::<Vec<_>>(),
            vec![10, 20, 30]
        );
        assert_eq!(
            index
                .find_objects_by_type(7)
                .into_iter()
                .map(|entry| entry.id().get())
                .collect::<Vec<_>>(),
            vec![10, 30]
        );
    }

    #[test]
    fn batch_resolution_preserves_request_order_across_fragments() {
        let first = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: 41,
                        data: Vec::new(),
                    }],
                )
                .unwrap(),
            ],
        };
        let second = Archive {
            objects: vec![
                ArchiveObject::new(
                    2,
                    vec![RawMessage {
                        type_: 42,
                        data: Vec::new(),
                    }],
                )
                .unwrap(),
            ],
        };
        let mut package = crate::IWorkPackage::new();
        package.replace_archive("Index/First.iwa", &first).unwrap();
        package
            .replace_archive("Index/Second.iwa", &second)
            .unwrap();
        let bundle = Bundle::from_bytes(&package.to_bytes().unwrap()).unwrap();
        let index = ObjectIndex::from_bundle(&bundle).unwrap();

        let resolved = index.resolve_objects(&bundle, &[2, 1]).unwrap();
        assert_eq!(
            resolved
                .into_iter()
                .map(|object| object.id().get())
                .collect::<Vec<_>>(),
            vec![2, 1]
        );
    }

    #[test]
    fn bundle_index_builds_reverse_references_in_archive_name_order() {
        let mut first = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 41,
                data: Vec::new(),
            }],
        )
        .unwrap();
        first.archive_info.message_infos[0].object_references = vec![3];
        let mut second = ArchiveObject::new(
            2,
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        second.archive_info.message_infos[0].object_references = vec![3];

        let mut package = crate::IWorkPackage::new();
        package
            .replace_archive(
                "Index/Z.iwa",
                &Archive {
                    objects: vec![second],
                },
            )
            .unwrap();
        package
            .replace_archive(
                "Index/A.iwa",
                &Archive {
                    objects: vec![first],
                },
            )
            .unwrap();
        let bundle = Bundle::from_bytes(&package.to_bytes().unwrap()).unwrap();
        let index = ObjectIndex::from_bundle(&bundle).unwrap();

        let target = ObjectId::try_from(3).unwrap();
        assert_eq!(
            index.dependents(target).unwrap().collect::<Vec<_>>(),
            vec![
                ObjectId::try_from(1).unwrap(),
                ObjectId::try_from(2).unwrap()
            ]
        );
    }

    #[test]
    fn rejects_object_ids_repeated_across_archives() {
        let object = |message_type| {
            ArchiveObject::new(
                7,
                vec![RawMessage {
                    type_: message_type,
                    data: Vec::new(),
                }],
            )
            .unwrap()
        };
        let mut package = crate::IWorkPackage::new();
        package
            .replace_archive(
                "Index/B.iwa",
                &Archive {
                    objects: vec![object(42)],
                },
            )
            .unwrap();
        package
            .replace_archive(
                "Index/A.iwa",
                &Archive {
                    objects: vec![object(43)],
                },
            )
            .unwrap();

        let bundle = Bundle::from_bytes(&package.to_bytes().unwrap()).unwrap();
        let error = ObjectIndex::from_bundle(&bundle).unwrap_err();
        assert!(matches!(
            error,
            Error::Archive(message)
                if message.contains("object 7")
                    && message.contains("Index/A.iwa")
                    && message.contains("Index/B.iwa")
        ));
    }

    #[test]
    fn typed_batch_resolution_rejects_unindexed_and_missing_objects() {
        let empty_package = crate::IWorkPackage::new().to_bytes().unwrap();
        let empty_bundle = Bundle::from_bytes(&empty_package).unwrap();
        let object_id = ObjectId::try_from(10).unwrap();

        let empty_index = ObjectIndex::new();
        let error = empty_index
            .resolve_many(&empty_bundle, &[object_id])
            .unwrap_err();
        assert!(matches!(error, Error::Archive(message) if message.contains("not present")));

        let object = ArchiveObject::new(
            object_id.get(),
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let mut index = ObjectIndex::new();
        index
            .parse_archive(
                "Index/Missing.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        let error = index.resolve_many(&empty_bundle, &[object_id]).unwrap_err();
        assert!(
            matches!(error, Error::Bundle(message) if message.contains("could not be resolved"))
        );
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
        assert_eq!(index.get_dependencies(10), Some(vec![20]));
        assert_eq!(index.get_dependencies(20), Some(vec![30]));
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
        assert_eq!(index.get_dependencies(10), Some(vec![20, 30]));
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
