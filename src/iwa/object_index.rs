//! Object Index for Cross-Referencing in iWork Documents
//!
//! iWork documents contain an object index that maps object IDs to their
//! locations in IWA files. This allows objects to reference each other
//! across different archive files.

use std::collections::HashMap;

use crate::iwa::archive::{Archive, ArchiveObject, RawMessage};
use crate::iwa::bundle::Bundle;
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
}

impl ObjectIndex {
    /// Create an empty object index
    pub fn new() -> Self {
        Self {
            entries: HashMap::new(),
            fragment_objects: HashMap::new(),
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
    fn parse_archive(&mut self, archive_name: &str, archive: &Archive) -> Result<()> {
        for object in &archive.objects {
            if let Some(identifier) = object.archive_info.identifier {
                // Determine object type from first message
                let object_type = object.messages.first()
                    .map(|msg| msg.type_)
                    .unwrap_or(0);

                let entry = ObjectIndexEntry {
                    id: identifier,
                    fragment_name: archive_name.to_string(),
                    data_offset: 0, // Would need to calculate actual offset
                    data_length: 0, // Would need to calculate actual length
                    object_type,
                };

                self.entries.insert(identifier, entry);
                self.fragment_objects.entry(archive_name.to_string())
                    .or_insert_with(Vec::new)
                    .push(identifier);
            }
        }
        Ok(())
    }

    /// Parse object references within an object's messages
    fn parse_object_references(&mut self, _object_id: u64, _object: &ArchiveObject) -> Result<()> {
        // This would parse the actual protobuf messages to find references
        // to other objects. For now, this is a placeholder.
        //
        // In a full implementation, this would:
        // 1. Decode the protobuf messages
        // 2. Look for fields containing object references
        // 3. Add entries to the reference graph
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
        self.entries.values()
            .filter(|entry| entry.object_type == object_type)
            .collect()
    }

    /// Resolve an object reference to get the actual object data
    pub fn resolve_object(&self, bundle: &Bundle, object_id: u64) -> Result<Option<ResolvedObject>> {
        let Some(entry) = self.get_entry(object_id) else {
            return Ok(None);
        };

        let Some(archive) = bundle.get_archive(&entry.fragment_name) else {
            return Err(Error::Bundle(format!("Archive {} not found", entry.fragment_name)));
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

/// Object reference graph for tracking dependencies
#[derive(Debug, Clone)]
pub struct ReferenceGraph {
    /// Map from object ID to objects that reference it
    incoming_refs: HashMap<u64, Vec<u64>>,
    /// Map from object ID to objects it references
    outgoing_refs: HashMap<u64, Vec<u64>>,
}

impl ReferenceGraph {
    /// Create an empty reference graph
    pub fn new() -> Self {
        Self {
            incoming_refs: HashMap::new(),
            outgoing_refs: HashMap::new(),
        }
    }

    /// Add a reference from source to target
    pub fn add_reference(&mut self, source_id: u64, target_id: u64) {
        self.outgoing_refs.entry(source_id)
            .or_insert_with(Vec::new)
            .push(target_id);
        self.incoming_refs.entry(target_id)
            .or_insert_with(Vec::new)
            .push(source_id);
    }

    /// Get objects that reference the given object
    pub fn get_incoming_refs(&self, object_id: u64) -> Option<&Vec<u64>> {
        self.incoming_refs.get(&object_id)
    }

    /// Get objects referenced by the given object
    pub fn get_outgoing_refs(&self, object_id: u64) -> Option<&Vec<u64>> {
        self.outgoing_refs.get(&object_id)
    }

    /// Get all object IDs in the graph
    pub fn all_objects(&self) -> std::collections::HashSet<u64> {
        let mut all = std::collections::HashSet::new();
        all.extend(self.incoming_refs.keys());
        all.extend(self.outgoing_refs.keys());
        all
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
    fn test_reference_graph() {
        let mut graph = ReferenceGraph::new();

        graph.add_reference(1, 2);
        graph.add_reference(1, 3);
        graph.add_reference(2, 3);

        assert_eq!(graph.get_outgoing_refs(1), Some(&vec![2, 3]));
        assert_eq!(graph.get_incoming_refs(3), Some(&vec![1, 2]));
        assert_eq!(graph.get_incoming_refs(1), None);
    }
}
