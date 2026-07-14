//! Object Index for Cross-Referencing in iWork Documents
//!
//! iWork documents contain an object index that maps object IDs to their
//! locations in IWA files. This allows objects to reference each other
//! across different archive files.

use std::collections::HashMap;

use crate::Result;
use crate::archive::{Archive, RawMessage};
use crate::bundle::Bundle;
use crate::ref_graph::ReferenceGraph;

mod references;
mod resolver;

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
    pub archive_info: crate::archive::ArchiveInfo,
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
