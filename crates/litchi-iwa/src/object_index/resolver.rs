//! Resolution of indexed object IDs into archive data.

use super::{ObjectIndex, ResolvedObject};
use crate::bundle::Bundle;
use crate::{Error, Result};

impl ObjectIndex {
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
}
