//! Object Index for Cross-Referencing in iWork Documents
//!
//! iWork documents contain an object index that maps object IDs to their
//! locations in IWA files. This allows objects to reference each other
//! across different archive files.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::bundle::Bundle;
use crate::{Error, Result};
use litchi_iwa_graph::{ObjectId, ObjectIdIter, ReferenceGraph};
use litchi_iwa_index::{
    ByteSpan, FragmentId, IndexBuilder, IndexError, ObjectIndex as NeutralObjectIndex, ObjectRecord,
};

mod reference_extraction;

/// Typed adapter metadata for one indexed object.
///
/// The neutral location snapshot owns the immutable object record and graph;
/// this record retains only the archive adapter metadata needed to resolve a
/// validated source position back into parsed IWA storage. Physical archive
/// names and source positions are deliberately not part of the public record.
#[derive(Debug, Clone)]
pub struct ObjectIndexEntry {
    /// Unique, validated object identifier.
    id: ObjectId,
    /// Which adapter-local fragment contains this object.
    fragment_id: FragmentId,
    /// Checked byte location within the fragment.
    span: ByteSpan,
    /// Which IWA file contains this object.
    fragment_name: Arc<str>,
    /// Position of the object within its parsed archive.
    ///
    /// This is an internal source-position hint. Resolution validates the
    /// identifier at the position and fails closed when a caller supplies a
    /// stale bundle with a different object order.
    object_position: usize,
    /// Native primary message type retained by the format adapter.
    object_type: u32,
}

impl ObjectIndexEntry {
    /// Return the validated object identity.
    pub const fn id(&self) -> ObjectId {
        self.id
    }

    /// Return the adapter-local fragment identity.
    pub const fn fragment_id(&self) -> FragmentId {
        self.fragment_id
    }

    /// Return the checked byte location within the fragment.
    pub const fn span(&self) -> ByteSpan {
        self.span
    }

    /// Return the native primary message type, or zero when the object has no
    /// messages.
    pub const fn object_type(&self) -> u32 {
        self.object_type
    }
}

#[derive(Debug, Clone)]
struct FragmentIndexEntry {
    id: FragmentId,
    name: Arc<str>,
}

#[derive(Debug, Clone)]
struct IndexSnapshot {
    locations: Arc<NeutralObjectIndex>,
    entries: Arc<[ObjectIndexEntry]>,
    fragments: Arc<[FragmentIndexEntry]>,
}

/// Object index that maps object IDs to their locations.
///
/// Archive decoding remains in this adapter, while immutable location and
/// graph storage is delegated to [`litchi_iwa_index::ObjectIndex`]. The
/// adapter retains only the metadata needed to resolve a validated source
/// position back into the already parsed archive.
#[derive(Debug, Clone)]
pub struct ObjectIndex {
    snapshot: Arc<IndexSnapshot>,
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
            snapshot: Arc::new(IndexSnapshot {
                locations: Arc::new(NeutralObjectIndex::default()),
                entries: Arc::default(),
                fragments: Arc::default(),
            }),
        }
    }

    /// Build object index from a bundle
    pub fn from_bundle(bundle: &Bundle) -> Result<Self> {
        let mut builder = IndexBuilder::new();
        let mut entries = Vec::new();
        let mut fragments = Vec::new();

        // Bundle traversal is already sorted at ingress, so assigning private
        // fragment ordinals here makes the neutral snapshot deterministic.
        for (position, (archive_name, archive)) in bundle.iter_archives().enumerate() {
            let fragment_id = fragment_id(position)?;
            let name: Arc<str> = Arc::from(archive_name);
            builder.add_fragment(fragment_id).map_err(index_error)?;
            fragments.push(FragmentIndexEntry {
                id: fragment_id,
                name: Arc::clone(&name),
            });
            append_archive(
                archive_name,
                archive,
                fragment_id,
                Arc::clone(&name),
                &mut builder,
                &mut entries,
            )?;
        }

        Ok(Self {
            snapshot: finish_snapshot(builder, entries, fragments)?,
        })
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
    #[cfg(test)]
    fn parse_archive(&mut self, archive_name: &str, archive: &Archive) -> Result<()> {
        if self
            .snapshot
            .fragments
            .iter()
            .any(|fragment| fragment.name.as_ref() == archive_name)
        {
            return Err(Error::Archive(format!(
                "archive {archive_name} occurs more than once in the object index"
            )));
        }

        let mut builder = IndexBuilder::new();
        for fragment in self.snapshot.fragments.iter() {
            builder.add_fragment(fragment.id).map_err(index_error)?;
        }
        for record in self.snapshot.locations.objects() {
            builder.add_object(*record).map_err(index_error)?;
        }
        let graph = self.snapshot.locations.reference_graph();
        for source in graph.iter_object_ids() {
            if let Some(targets) = graph.outgoing(source) {
                for target in targets {
                    add_reference(&mut builder, source, target)?;
                }
            }
        }

        let fragment_id = fragment_id(self.snapshot.fragments.len())?;
        builder.add_fragment(fragment_id).map_err(index_error)?;
        let name: Arc<str> = Arc::from(archive_name);
        let mut entries = self.snapshot.entries.to_vec();
        let mut fragments = self.snapshot.fragments.to_vec();
        fragments.push(FragmentIndexEntry {
            id: fragment_id,
            name: Arc::clone(&name),
        });
        append_archive(
            archive_name,
            archive,
            fragment_id,
            name,
            &mut builder,
            &mut entries,
        )?;
        self.snapshot = finish_snapshot(builder, entries, fragments)?;
        Ok(())
    }

    /// Get an object entry through the validated identity API.
    pub fn entry(&self, object_id: ObjectId) -> Option<&ObjectIndexEntry> {
        self.snapshot
            .entries
            .binary_search_by_key(&object_id, ObjectIndexEntry::id)
            .ok()
            .and_then(|position| self.snapshot.entries.get(position))
    }

    /// Borrow all validated object identities in deterministic numeric order.
    ///
    /// The index validates identities while it is built and stores this order
    /// as compact immutable neutral records, so traversal does not allocate or
    /// depend on randomized hash-map order.
    pub fn iter_object_ids(&self) -> impl Iterator<Item = ObjectId> + '_ {
        self.snapshot.locations.objects().map(ObjectRecord::id)
    }

    /// Get all indexed object identities in deterministic numeric order.
    ///
    /// This is an owned convenience collection over [`Self::iter_object_ids`].
    /// The index invariants make the operation infallible; callers that only
    /// need to inspect the catalog should prefer the borrowed iterator.
    pub fn object_ids(&self) -> Vec<ObjectId> {
        self.iter_object_ids().collect()
    }

    /// Borrow typed object identities for one fragment in deterministic ID order.
    ///
    /// The identities are validated while the index is built, so this view
    /// performs no allocation or repeated wire-boundary conversion.
    pub fn fragment_object_ids(&self, fragment_name: &str) -> Option<&[ObjectId]> {
        let fragment = self
            .snapshot
            .fragments
            .binary_search_by(|fragment| fragment.name.as_ref().cmp(fragment_name))
            .ok()
            .and_then(|position| self.snapshot.fragments.get(position))?;
        self.snapshot.locations.fragment_object_ids(fragment.id)
    }

    /// Borrow all entries in deterministic numeric object-ID order.
    pub fn iter_entries(&self) -> impl Iterator<Item = &ObjectIndexEntry> {
        self.snapshot.entries.iter()
    }

    /// Get entries of one type in deterministic numeric object-ID order.
    pub fn iter_entries_by_type(
        &self,
        object_type: u32,
    ) -> impl Iterator<Item = &ObjectIndexEntry> {
        self.iter_entries()
            .filter(move |entry| entry.object_type() == object_type)
    }

    /// Collect all entries in deterministic numeric object-ID order.
    pub fn all_entries(&self) -> Vec<&ObjectIndexEntry> {
        self.iter_entries().collect()
    }

    /// Find objects by type in deterministic numeric object-ID order.
    pub fn find_objects_by_type(&self, object_type: u32) -> Vec<&ObjectIndexEntry> {
        self.iter_entries_by_type(object_type).collect()
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
    pub fn reference_graph(&self) -> &litchi_iwa_graph::ReferenceGraphSnapshot {
        self.snapshot.locations.reference_graph()
    }

    /// Get typed dependencies without exposing raw sentinel IDs.
    pub fn dependencies(&self, object_id: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.snapshot.locations.outgoing(object_id)
    }

    /// Get typed dependents without exposing raw sentinel IDs.
    pub fn dependents(&self, object_id: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.snapshot.locations.incoming(object_id)
    }

    /// Check for a cycle through the validated identity API.
    pub fn has_cycle_from(&self, object_id: ObjectId) -> bool {
        self.snapshot.locations.has_cycle(object_id)
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
    ///
    /// Get typed transitive dependencies, including the starting object.
    pub fn reachable_from(&self, object_id: ObjectId) -> Vec<ObjectId> {
        self.snapshot.locations.reachable(object_id)
    }

    /// Borrow a protobuf wire identifier for crate-internal readers.
    pub(crate) fn resolve_ref_id<'a>(
        &self,
        bundle: &'a Bundle,
        object_id: u64,
    ) -> Result<Option<ResolvedObjectRef<'a>>> {
        let Some(object_id) = ObjectId::new(object_id) else {
            return Ok(None);
        };
        self.resolve_ref(bundle, object_id)
    }

    /// Borrow an indexed object directly from the supplied bundle.
    ///
    /// The returned view borrows the bundle's immutable archive storage, so
    /// resolving an object does not clone its archive metadata or message
    /// payloads. The view cannot outlive `bundle`; use [`Self::resolve`] when
    /// an owned result must be retained after the bundle is dropped.
    pub fn resolve_ref<'a>(
        &self,
        bundle: &'a Bundle,
        object_id: ObjectId,
    ) -> Result<Option<ResolvedObjectRef<'a>>> {
        let Some(entry) = self.entry(object_id) else {
            return Ok(None);
        };

        let Some(record) = self.snapshot.locations.object(object_id) else {
            return Err(Error::Archive(format!(
                "object {} is missing from the neutral index",
                object_id.get()
            )));
        };
        if record.fragment() != entry.fragment_id || record.span() != entry.span {
            return Err(Error::Archive(format!(
                "object {} has inconsistent neutral location metadata",
                object_id.get()
            )));
        }

        let Some(archive) = bundle.get_archive(entry.fragment_name.as_ref()) else {
            return Err(Error::Bundle(format!(
                "Archive {} not found",
                entry.fragment_name
            )));
        };

        let object = indexed_object(archive, entry, object_id)?;

        Ok(Some(ResolvedObjectRef {
            id: object_id,
            archive_info: &object.archive_info,
            messages: &object.messages,
        }))
    }

    /// Borrow every indexed object in deterministic numeric-ID order.
    ///
    /// The iterator performs no collection allocation. Each item validates
    /// the indexed source position and returns a borrowed view tied to the
    /// supplied immutable bundle. Use [`Self::resolve_many_refs`] when an
    /// owned collection of views is required.
    pub fn iter_refs<'a>(
        &'a self,
        bundle: &'a Bundle,
    ) -> impl Iterator<Item = Result<ResolvedObjectRef<'a>>> + 'a {
        self.iter_entries().map(move |entry| {
            self.resolve_ref(bundle, entry.id())?.ok_or_else(|| {
                Error::Bundle(format!(
                    "object {} could not be resolved from the bundle",
                    entry.id().get()
                ))
            })
        })
    }

    /// Resolve an object through the validated identity API.
    pub fn resolve(&self, bundle: &Bundle, object_id: ObjectId) -> Result<Option<ResolvedObject>> {
        self.resolve_ref(bundle, object_id)
            .map(|object| object.map(ResolvedObjectRef::into_owned))
    }
    /// Borrow multiple indexed objects in the caller's request order.
    ///
    /// Archive lookups are grouped by fragment, while the returned views
    /// borrow the original bundle and retain no duplicate payload allocation.
    /// Duplicate typed IDs are rejected just like [`Self::resolve_many`].
    pub fn resolve_many_refs<'a>(
        &self,
        bundle: &'a Bundle,
        object_ids: &[ObjectId],
    ) -> Result<Vec<ResolvedObjectRef<'a>>> {
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

        let resolved = self.resolve_many_refs_inner(bundle, object_ids)?;
        let resolved_ids: HashSet<_> = resolved.iter().map(ResolvedObjectRef::id).collect();
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

    fn resolve_many_refs_inner<'a>(
        &self,
        bundle: &'a Bundle,
        object_ids: &[ObjectId],
    ) -> Result<Vec<ResolvedObjectRef<'a>>> {
        // Group object IDs by their archive to minimize archive lookups
        let mut objects_by_archive: std::collections::HashMap<&str, HashSet<ObjectId>> =
            std::collections::HashMap::new();

        for &object_id in object_ids {
            if let Some(entry) = self.entry(object_id) {
                objects_by_archive
                    .entry(entry.fragment_name.as_ref())
                    .or_default()
                    .insert(object_id);
            }
        }

        let mut resolved_by_id = HashMap::with_capacity(object_ids.len());

        // Resolve objects archive by archive. The indexed source position
        // avoids rescanning each archive for sparse batches; the helper keeps
        // the compatibility fallback for a bundle with a different order.
        for (archive_name, ids) in objects_by_archive {
            if let Some(archive) = bundle.get_archive(archive_name) {
                for object_id in ids {
                    let Some(entry) = self.entry(object_id) else {
                        continue;
                    };
                    let object = indexed_object(archive, entry, object_id)?;
                    let resolved = ResolvedObjectRef {
                        id: object_id,
                        archive_info: &object.archive_info,
                        messages: &object.messages,
                    };
                    if resolved_by_id.insert(object_id, resolved).is_some() {
                        return Err(Error::Archive(format!(
                            "object {} occurs in more than one archive",
                            object_id.get()
                        )));
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
        self.resolve_many_refs(bundle, object_ids).map(|objects| {
            objects
                .into_iter()
                .map(ResolvedObjectRef::into_owned)
                .collect()
        })
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

    /// Check for an indexed object through the validated identity API.
    pub fn contains(&self, object_id: ObjectId) -> bool {
        self.snapshot.locations.object(object_id).is_some()
    }

    /// Get the total number of indexed objects
    pub fn object_count(&self) -> usize {
        self.snapshot.locations.len()
    }

    /// Get the number of fragments (IWA files) in the index
    pub fn fragment_count(&self) -> usize {
        self.snapshot.locations.fragment_count()
    }

    /// Get statistics about the object index
    pub fn stats(&self) -> ObjectIndexStats {
        let total_objects = self.snapshot.locations.len();
        let total_fragments = self.snapshot.locations.fragment_count();
        let total_references = self.snapshot.locations.reference_graph().edge_count();
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

fn fragment_id(position: usize) -> Result<FragmentId> {
    let ordinal = position
        .checked_add(1)
        .and_then(|value| u32::try_from(value).ok())
        .ok_or_else(|| Error::Archive("IWA fragment catalog exceeds u32 capacity".to_owned()))?;
    FragmentId::try_from(ordinal)
        .map_err(|error| Error::Archive(format!("invalid IWA fragment ordinal: {error}")))
}

fn append_archive(
    archive_name: &str,
    archive: &Archive,
    fragment_id: FragmentId,
    fragment_name: Arc<str>,
    builder: &mut IndexBuilder,
    entries: &mut Vec<ObjectIndexEntry>,
) -> Result<()> {
    for (object_position, object) in archive.objects.iter().enumerate() {
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

        let object_type = object.messages.first().map_or(0, |message| message.type_);
        let span = ByteSpan::new(object.data_offset, object.data_length).map_err(|error| {
            Error::Archive(format!(
                "archive {archive_name} object {identifier} has an invalid byte span: {error}"
            ))
        })?;
        if let Err(error) = builder.add_object(ObjectRecord::new(object_id, fragment_id, span)) {
            if matches!(error, IndexError::DuplicateObject(_))
                && let Some(existing) = entries.iter().find(|entry| entry.id() == object_id)
            {
                return Err(Error::Archive(format!(
                    "object {identifier} occurs in archives {} and {archive_name}",
                    existing.fragment_name
                )));
            }
            return Err(index_error(error));
        }
        entries.push(ObjectIndexEntry {
            id: object_id,
            fragment_id,
            span,
            fragment_name: Arc::clone(&fragment_name),
            object_position,
            object_type,
        });

        // MessageInfo is the authoritative, application-independent
        // reference index emitted by iWork for every payload.
        let mut graph = ReferenceGraph::new();
        let mut has_indexed_references = false;
        for message_info in &object.archive_info.message_infos {
            has_indexed_references |= !message_info.object_references.is_empty();
            for &reference in &message_info.object_references {
                if let Some(target_id) = ObjectId::new(reference) {
                    graph.add_object_reference(object_id, target_id);
                }
            }
        }

        // Some old archives omit MessageInfo references. Decode only
        // unambiguous high-numbered payloads as a compatibility fallback;
        // low message types overlap between Numbers and Keynote.
        if !has_indexed_references && object_type >= 2000 {
            reference_extraction::extract(object_id, object, &mut graph)?;
        }

        let graph = graph.snapshot();
        if let Some(targets) = graph.outgoing(object_id) {
            for target in targets {
                add_reference(builder, object_id, target)?;
            }
        }
    }
    Ok(())
}

fn add_reference(builder: &mut IndexBuilder, source: ObjectId, target: ObjectId) -> Result<()> {
    builder.add_reference(source, target).map_err(index_error)
}

fn finish_snapshot(
    builder: IndexBuilder,
    mut entries: Vec<ObjectIndexEntry>,
    mut fragments: Vec<FragmentIndexEntry>,
) -> Result<Arc<IndexSnapshot>> {
    entries.sort_unstable_by_key(ObjectIndexEntry::id);
    fragments.sort_unstable_by(|left, right| left.name.cmp(&right.name));
    let locations = builder.build_allow_missing_targets().map_err(index_error)?;
    Ok(Arc::new(IndexSnapshot {
        locations: Arc::new(locations),
        entries: Arc::from(entries.into_boxed_slice()),
        fragments: Arc::from(fragments.into_boxed_slice()),
    }))
}

fn index_error(error: IndexError) -> Error {
    Error::Archive(format!("object index construction failed: {error}"))
}

/// Resolve an indexed object by its source position, validating the identity
/// before returning it.
///
/// The parsed archive position is the index's authoritative lookup key. A
/// separately reordered or truncated archive is a stale snapshot, not a
/// reason to perform an unbounded linear scan, so it fails closed with a
/// contextual archive error.
fn indexed_object<'a>(
    archive: &'a Archive,
    entry: &ObjectIndexEntry,
    object_id: ObjectId,
) -> Result<&'a ArchiveObject> {
    let object = archive.objects.get(entry.object_position).ok_or_else(|| {
        Error::Archive(format!(
            "object {} in archive {} has stale source position {}",
            object_id.get(),
            entry.fragment_name,
            entry.object_position
        ))
    })?;

    if object.archive_info.identifier != Some(object_id.get()) {
        return Err(Error::Archive(format!(
            "object {} in archive {} has stale source position {} (found identifier {:?})",
            object_id.get(),
            entry.fragment_name,
            entry.object_position,
            object.archive_info.identifier
        )));
    }

    Ok(object)
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

/// A borrowed view of an indexed object and its immutable payloads.
///
/// The view is tied to the [`crate::raw::bundle::Bundle`] used for resolution. It is the
/// allocation-free read path for traversal and extraction; callers that need
/// an owned value can consume it with [`Self::into_owned`].
#[derive(Debug, Clone, Copy)]
pub struct ResolvedObjectRef<'a> {
    /// Validated object identifier.
    id: ObjectId,
    /// Borrowed archive information.
    pub archive_info: &'a crate::archive::ArchiveInfo,
    /// Borrowed raw message data.
    pub messages: &'a [RawMessage],
}

impl ResolvedObjectRef<'_> {
    /// Return the validated object identity.
    pub const fn id(&self) -> ObjectId {
        self.id
    }

    /// Return the validated object identity, if the compatibility payload is
    /// non-null.
    pub const fn object_id(&self) -> Option<ObjectId> {
        Some(self.id)
    }

    /// Get the primary message type without allocating.
    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|message| message.type_)
    }

    /// Iterate over message types without cloning the message payloads.
    pub fn message_types(&self) -> impl Iterator<Item = u32> + '_ {
        self.messages.iter().map(|message| message.type_)
    }

    /// Clone the borrowed payloads into the legacy owned representation.
    pub fn into_owned(self) -> ResolvedObject {
        ResolvedObject {
            id: self.id,
            archive_info: self.archive_info.clone(),
            messages: self.messages.to_vec(),
        }
    }
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
        assert!(index.snapshot.locations.is_empty());
        assert!(index.snapshot.entries.is_empty());
        assert!(index.snapshot.fragments.is_empty());
    }

    #[test]
    fn object_index_clones_share_immutable_snapshot_until_indexing_mutates() {
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
        assert!(Arc::ptr_eq(&index.snapshot, &snapshot.snapshot));

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
        assert!(!Arc::ptr_eq(&index.snapshot, &edited.snapshot));
    }

    #[test]
    fn object_indexes_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<ObjectIndex>();
    }

    #[test]
    fn test_object_index_entry() {
        let fragment_id = FragmentId::try_from(1).unwrap();
        let span = ByteSpan::new(100, 200).unwrap();
        let entry = ObjectIndexEntry {
            id: ObjectId::try_from(123).unwrap(),
            fragment_id,
            span,
            fragment_name: Arc::from("Document.iwa"),
            object_position: 0,
            object_type: 42,
        };

        assert_eq!(entry.id().get(), 123);
        assert_eq!(entry.fragment_id(), fragment_id);
        assert_eq!(entry.span(), span);
        assert_eq!(entry.object_type(), 42);
    }

    #[test]
    fn indexed_object_positions_fail_closed_for_stale_archives() {
        let object = |identifier| {
            ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: 42,
                    data: Vec::new(),
                }],
            )
            .unwrap()
        };
        let original = Archive {
            objects: vec![object(10), object(20)],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Test.iwa", &original).unwrap();

        let object_id = ObjectId::try_from(10).unwrap();
        let entry = index.entry(object_id).unwrap();
        assert_eq!(
            indexed_object(&original, entry, object_id)
                .unwrap()
                .archive_info
                .identifier,
            Some(10)
        );

        let reordered = Archive {
            objects: vec![object(20), object(10)],
        };
        let error = indexed_object(&reordered, entry, object_id).unwrap_err();
        assert!(matches!(
            error,
            Error::Archive(message) if message.contains("stale source position")
        ));

        let truncated = Archive {
            objects: vec![object(10)],
        };
        let object_id = ObjectId::try_from(20).unwrap();
        let entry = index.entry(object_id).unwrap();
        let error = indexed_object(&truncated, entry, object_id).unwrap_err();
        assert!(matches!(
            error,
            Error::Archive(message) if message.contains("stale source position")
        ));
    }

    #[test]
    fn indexed_object_reads_scale_without_linear_rescans() {
        const OBJECT_COUNT: u64 = 4096;

        let objects = (1..=OBJECT_COUNT)
            .map(|identifier| {
                ArchiveObject::new(
                    identifier,
                    vec![RawMessage {
                        type_: 42,
                        data: Vec::new(),
                    }],
                )
                .unwrap()
            })
            .collect();
        let archive = Archive { objects };
        let mut index = ObjectIndex::new();
        index
            .parse_archive("Index/Benchmark.iwa", &archive)
            .unwrap();

        // Benchmark-shaped consumption: resolve the whole ordered catalog
        // through borrowed references without collecting a second object list.
        let (count, last_id) = index
            .iter_entries()
            .try_fold((0usize, 0u64), |(count, previous_id), entry| {
                let object = indexed_object(&archive, entry, entry.id())?;
                let id = entry.id().get();
                assert!(id > previous_id);
                assert_eq!(object.archive_info.identifier, Some(id));
                Ok::<(usize, u64), Error>((count + 1, id))
            })
            .unwrap();

        assert_eq!(count, OBJECT_COUNT as usize);
        assert_eq!(last_id, OBJECT_COUNT);
    }

    #[test]
    fn test_object_index_with_reference_graph() {
        let index = ObjectIndex::new();
        let object_id = ObjectId::try_from(1).unwrap();

        assert!(index.reference_graph().is_empty());
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

        let source = ObjectId::try_from(10).unwrap();
        let target = ObjectId::try_from(20).unwrap();
        assert_eq!(
            index
                .dependencies(source)
                .map(|references| references.collect::<Vec<_>>()),
            Some(vec![target, ObjectId::try_from(30).unwrap()])
        );
        assert_eq!(
            index
                .dependents(target)
                .map(|references| references.collect::<Vec<_>>()),
            Some(vec![source])
        );
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

        assert!(
            index
                .dependencies(ObjectId::try_from(10).unwrap())
                .is_none()
        );
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

        assert_eq!(index.entry(source).map(ObjectIndexEntry::id), Some(source));
        let entry = index.entry(source).unwrap();
        assert_eq!(entry.fragment_id(), FragmentId::try_from(1).unwrap());
        assert_eq!(entry.span(), ByteSpan::new(0, 0).unwrap());
        assert_eq!(entry.object_type(), 42);
        assert_eq!(index.object_ids(), vec![source]);
        assert_eq!(index.iter_object_ids().collect::<Vec<_>>(), vec![source]);
        assert_eq!(
            index.fragment_object_ids("Index/Test.iwa"),
            Some([source].as_slice())
        );
        assert_eq!(index.fragment_object_ids("missing.iwa"), None);
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
    fn typed_object_queries_are_deterministically_ordered() {
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

        assert_eq!(
            index
                .iter_object_ids()
                .map(ObjectId::get)
                .collect::<Vec<_>>(),
            vec![10, 20, 30]
        );
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
                .iter_entries()
                .map(|entry| entry.id().get())
                .collect::<Vec<_>>(),
            vec![10, 20, 30]
        );
        assert_eq!(
            index
                .iter_entries_by_type(7)
                .map(|entry| entry.id().get())
                .collect::<Vec<_>>(),
            vec![10, 30]
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

        let typed_ids = [
            ObjectId::try_from(2).unwrap(),
            ObjectId::try_from(1).unwrap(),
        ];
        let borrowed = index.resolve_many_refs(&bundle, &typed_ids).unwrap();
        assert_eq!(
            borrowed
                .iter()
                .map(ResolvedObjectRef::id)
                .collect::<Vec<_>>(),
            typed_ids
        );
        assert_eq!(borrowed[0].primary_message_type(), Some(42));
        assert_eq!(borrowed[0].message_types().collect::<Vec<_>>(), vec![42]);
        assert_eq!(borrowed[0].messages[0].data, Vec::<u8>::new());

        let streamed = index
            .iter_refs(&bundle)
            .collect::<Result<Vec<_>>>()
            .unwrap();
        assert_eq!(
            streamed
                .iter()
                .map(|object| object.id())
                .collect::<Vec<_>>(),
            [
                ObjectId::try_from(1).unwrap(),
                ObjectId::try_from(2).unwrap()
            ]
        );

        let resolved = index
            .resolve_many(
                &bundle,
                &[
                    ObjectId::try_from(2).unwrap(),
                    ObjectId::try_from(1).unwrap(),
                ],
            )
            .unwrap();
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
        assert_eq!(
            index
                .dependencies(ObjectId::try_from(10).unwrap())
                .map(|references| references.collect::<Vec<_>>()),
            Some(vec![ObjectId::try_from(20).unwrap()])
        );
        assert_eq!(
            index
                .dependencies(ObjectId::try_from(20).unwrap())
                .map(|references| references.collect::<Vec<_>>()),
            Some(vec![ObjectId::try_from(30).unwrap()])
        );
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
        assert_eq!(
            index
                .dependencies(ObjectId::try_from(10).unwrap())
                .map(|references| references.collect::<Vec<_>>()),
            Some(vec![
                ObjectId::try_from(20).unwrap(),
                ObjectId::try_from(30).unwrap()
            ])
        );
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

        let document_dependencies = index
            .dependencies(ObjectId::try_from(1).unwrap())
            .unwrap()
            .collect::<Vec<_>>();
        for identifier in [42, 43, 44, 45] {
            assert!(document_dependencies.contains(&ObjectId::try_from(identifier).unwrap()));
        }
        let section_dependencies = index
            .dependencies(ObjectId::try_from(43).unwrap())
            .unwrap()
            .collect::<Vec<_>>();
        for identifier in [50, 51, 52, 53] {
            assert!(section_dependencies.contains(&ObjectId::try_from(identifier).unwrap()));
        }
        let template_dependencies = index
            .dependencies(ObjectId::try_from(50).unwrap())
            .unwrap()
            .collect::<Vec<_>>();
        assert_eq!(
            template_dependencies,
            [
                ObjectId::try_from(60).unwrap(),
                ObjectId::try_from(61).unwrap(),
                ObjectId::try_from(62).unwrap()
            ]
        );
    }
}
