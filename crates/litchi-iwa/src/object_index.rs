//! Object Index for Cross-Referencing in iWork Documents
//!
//! iWork documents contain an object index that maps object IDs to their
//! locations in IWA files. This allows objects to reference each other
//! across different archive files.

use std::collections::HashMap;
#[cfg(test)]
use std::collections::HashSet;
use std::sync::Arc;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::bundle::Bundle;
use crate::{Error, Result};
#[cfg(test)]
use litchi_iwa_index::Reference;
use litchi_iwa_index::{
    ByteSpan, FragmentId, IndexBuilder, IndexError, ObjectId, ObjectIndex as NeutralObjectIndex,
    ObjectRecord,
};

mod reference_extraction;

/// Adapter-only source metadata for one neutral object record.
///
/// The metadata is stored in the same object-ID order as the neutral index.
/// It deliberately contains no object identity, fragment identity, or byte
/// span: those values have one owner in [`litchi_iwa_index::ObjectRecord`].
#[derive(Debug, Clone, Copy)]
struct ArchiveObjectMetadata {
    /// Position of the object within its parsed archive.
    source_position: ArchiveObjectPosition,
    /// Native primary message type retained by the format adapter.
    object_type: u32,
}

/// A private, typed source position for an already-parsed native archive.
///
/// This is intentionally distinct from the neutral index's object-slice
/// position and from [`ByteSpan`]. It is never exposed through the public
/// object-index API.
#[derive(Debug, Clone, Copy)]
struct ArchiveObjectPosition(usize);

impl ArchiveObjectPosition {
    const fn new(position: usize) -> Self {
        Self(position)
    }

    const fn get(self) -> usize {
        self.0
    }
}

/// Temporary adapter metadata collected while native archives are traversed.
///
/// The object identity is used only to align the sidecar with the sorted
/// neutral records during snapshot construction, then is dropped.
#[derive(Debug, Clone, Copy)]
struct PendingObjectMetadata {
    id: ObjectId,
    source_position: ArchiveObjectPosition,
    object_type: u32,
}

/// Typed borrowed metadata for one indexed object.
///
/// The neutral location snapshot owns the immutable object record and graph;
/// this view borrows that record and combines it with only the archive adapter
/// metadata needed to resolve a validated source position. No neutral
/// identity, fragment, or byte-span value is duplicated in the adapter.
#[derive(Debug, Clone, Copy)]
pub struct ObjectIndexEntry<'a> {
    record: &'a ObjectRecord,
    metadata: &'a ArchiveObjectMetadata,
}

impl ObjectIndexEntry<'_> {
    /// Return the validated object identity.
    pub const fn id(&self) -> ObjectId {
        self.record.id()
    }

    /// Return the adapter-local fragment identity.
    pub const fn fragment_id(&self) -> FragmentId {
        self.record.fragment()
    }

    /// Return the checked byte location within the fragment.
    pub const fn span(&self) -> ByteSpan {
        self.record.span()
    }

    /// Return the native primary message type, or zero when the object has no
    /// messages.
    pub const fn object_type(&self) -> u32 {
        self.metadata.object_type
    }

    const fn source_position(&self) -> ArchiveObjectPosition {
        self.metadata.source_position
    }
}

#[cfg(test)]
#[derive(Debug, Clone, Copy)]
struct BatchObject<'a> {
    request_position: usize,
    entry: ObjectIndexEntry<'a>,
}

#[derive(Debug, Clone)]
struct FragmentIndexEntry {
    id: FragmentId,
    name: Arc<str>,
}

#[derive(Clone)]
struct IndexSnapshot {
    locations: Arc<NeutralObjectIndex>,
    metadata: Arc<[ArchiveObjectMetadata]>,
    fragments: Arc<[FragmentIndexEntry]>,
    /// Fragment names keyed by their adapter-local identity.
    ///
    /// `fragments` remains name-sorted for the public name lookup, while this
    /// sidecar keeps source resolution independent of the number of package
    /// fragments. The names themselves stay shared through `Arc<str>`.
    fragment_names: Arc<HashMap<FragmentId, Arc<str>>>,
}

impl std::fmt::Debug for IndexSnapshot {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        // Keep the private ID lookup sidecar out of diagnostics: HashMap's
        // iteration order is randomized, while the published archive and
        // object views are intentionally deterministic.
        formatter
            .debug_struct("IndexSnapshot")
            .field("locations", &self.locations)
            .field("metadata", &self.metadata)
            .field("fragments", &self.fragments)
            .finish()
    }
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
                metadata: Arc::default(),
                fragments: Arc::default(),
                fragment_names: Arc::default(),
            }),
        }
    }

    /// Build object index from a bundle
    pub fn from_bundle(bundle: &Bundle) -> Result<Self> {
        let cardinality = preflight_bundle(bundle)?;
        let mut builder = IndexBuilder::new();
        let mut metadata = Vec::new();
        reserve_vec_exact(
            &mut metadata,
            cardinality.objects,
            "object index pending metadata",
        )?;
        let mut fragments = Vec::new();
        reserve_vec_exact(
            &mut fragments,
            cardinality.fragments,
            "object index fragment catalog",
        )?;

        // IndexBuilder keeps its own validation catalogs, but the adapter
        // also needs the existing object's fragment to report a useful,
        // deterministic cross-archive duplicate error. Pre-sizing this
        // private map makes duplicate diagnostics O(1) even for a hostile
        // package containing many repeated IDs.
        let mut object_fragments = HashMap::new();
        reserve_map(
            &mut object_fragments,
            cardinality.objects,
            "object index object-fragment catalog",
        )?;
        let mut fragment_names = HashMap::new();
        reserve_map(
            &mut fragment_names,
            cardinality.fragments,
            "object index fragment-name catalog",
        )?;

        // Bundle traversal is already sorted at ingress, so assigning private
        // fragment ordinals here makes the neutral snapshot deterministic.
        for (position, (archive_name, archive)) in bundle.iter_archives().enumerate() {
            let fragment_id = fragment_id(position)?;
            let name: Arc<str> = Arc::from(archive_name);
            builder.add_fragment(fragment_id).map_err(index_error)?;
            fragment_names.insert(fragment_id, Arc::clone(&name));
            fragments.push(FragmentIndexEntry {
                id: fragment_id,
                name: Arc::clone(&name),
            });
            append_archive(
                archive_name,
                archive,
                fragment_id,
                &mut builder,
                &mut metadata,
                &fragment_names,
                &mut object_fragments,
            )?;
        }

        Ok(Self {
            snapshot: finish_snapshot(builder, metadata, fragments)?,
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
            .binary_search_by(|fragment| fragment.name.as_ref().cmp(archive_name))
            .is_ok()
        {
            return Err(Error::Archive(format!(
                "archive {archive_name} occurs more than once in the object index"
            )));
        }

        let object_count = self
            .snapshot
            .locations
            .len()
            .checked_add(archive.objects.len())
            .ok_or_else(|| {
                Error::Archive("object index object cardinality overflows usize".to_owned())
            })?;
        let fragment_count = self
            .snapshot
            .fragments
            .len()
            .checked_add(1)
            .ok_or_else(|| {
                Error::Archive("object index fragment cardinality overflows usize".to_owned())
            })?;
        if fragment_count > u32::MAX as usize {
            return Err(Error::Archive(
                "IWA fragment catalog exceeds u32 capacity".to_owned(),
            ));
        }

        let mut builder = IndexBuilder::new();
        for fragment in self.snapshot.fragments.iter() {
            builder.add_fragment(fragment.id).map_err(index_error)?;
        }
        for record in self.snapshot.locations.objects() {
            builder.add_object(*record).map_err(index_error)?;
        }
        for reference in self.snapshot.locations.references() {
            builder
                .add_reference(reference.source(), reference.target())
                .map_err(index_error)?;
        }

        let fragment_id = fragment_id(self.snapshot.fragments.len())?;
        builder.add_fragment(fragment_id).map_err(index_error)?;
        let name: Arc<str> = Arc::from(archive_name);
        let mut metadata = Vec::new();
        reserve_vec_exact(&mut metadata, object_count, "object index pending metadata")?;
        for (record, object_metadata) in self
            .snapshot
            .locations
            .objects()
            .zip(self.snapshot.metadata.iter())
        {
            metadata.push(PendingObjectMetadata {
                id: record.id(),
                source_position: object_metadata.source_position,
                object_type: object_metadata.object_type,
            });
        }
        let mut fragments = Vec::new();
        reserve_vec_exact(
            &mut fragments,
            fragment_count,
            "object index fragment catalog",
        )?;
        fragments.extend(self.snapshot.fragments.iter().cloned());
        fragments.push(FragmentIndexEntry {
            id: fragment_id,
            name: Arc::clone(&name),
        });

        let mut object_fragments = HashMap::new();
        reserve_map(
            &mut object_fragments,
            object_count,
            "object index object-fragment catalog",
        )?;
        for record in self.snapshot.locations.objects() {
            object_fragments.insert(record.id(), record.fragment());
        }

        let mut fragment_names = HashMap::new();
        reserve_map(
            &mut fragment_names,
            fragment_count,
            "object index fragment-name catalog",
        )?;
        for fragment in self.snapshot.fragments.iter() {
            fragment_names.insert(fragment.id, Arc::clone(&fragment.name));
        }
        fragment_names.insert(fragment_id, Arc::clone(&name));

        append_archive(
            archive_name,
            archive,
            fragment_id,
            &mut builder,
            &mut metadata,
            &fragment_names,
            &mut object_fragments,
        )?;
        self.snapshot = finish_snapshot(builder, metadata, fragments)?;
        Ok(())
    }

    /// Get an object entry through the validated identity API.
    pub fn entry(&self, object_id: ObjectId) -> Option<ObjectIndexEntry<'_>> {
        let (position, record) = self.snapshot.locations.object_with_position(object_id)?;
        self.snapshot
            .metadata
            .get(position)
            .map(|metadata| ObjectIndexEntry { record, metadata })
    }

    /// Borrow all validated object identities in deterministic numeric order.
    ///
    /// The index validates identities while it is built and stores this order
    /// as compact immutable neutral records, so traversal does not allocate or
    /// depend on randomized hash-map order.
    #[cfg(test)]
    pub fn iter_object_ids(&self) -> impl Iterator<Item = ObjectId> + '_ {
        self.snapshot.locations.objects().map(ObjectRecord::id)
    }

    /// Get all indexed object identities in deterministic numeric order.
    ///
    /// This is an owned convenience collection over [`Self::iter_object_ids`].
    /// The index invariants make the operation infallible; callers that only
    /// need to inspect the catalog should prefer the borrowed iterator.
    #[cfg(test)]
    pub fn object_ids(&self) -> Vec<ObjectId> {
        self.iter_object_ids().collect()
    }

    /// Borrow typed object identities for one fragment in deterministic ID order.
    ///
    /// The identities are validated while the index is built, so this view
    /// performs no allocation or repeated wire-boundary conversion.
    #[cfg(test)]
    pub fn fragment_object_ids(&self, fragment_name: &str) -> Option<&[ObjectId]> {
        let fragment = self
            .snapshot
            .fragments
            .binary_search_by(|fragment| fragment.name.as_ref().cmp(fragment_name))
            .ok()
            .and_then(|position| self.snapshot.fragments.get(position))?;
        self.snapshot.locations.fragment_object_ids(fragment.id)
    }

    fn fragment_name(&self, fragment_id: FragmentId) -> Result<&str> {
        self.snapshot
            .fragment_names
            .get(&fragment_id)
            .map(Arc::as_ref)
            .ok_or_else(|| {
                Error::Archive(format!(
                    "object index references unregistered fragment {fragment_id:?}"
                ))
            })
    }

    /// Borrow all entries in deterministic numeric object-ID order.
    pub fn iter_entries(&self) -> impl Iterator<Item = ObjectIndexEntry<'_>> {
        self.snapshot
            .locations
            .objects()
            .zip(self.snapshot.metadata.iter())
            .map(|(record, metadata)| ObjectIndexEntry { record, metadata })
    }

    /// Get entries of one type in deterministic numeric object-ID order.
    pub fn iter_entries_by_type(
        &self,
        object_type: u32,
    ) -> impl Iterator<Item = ObjectIndexEntry<'_>> {
        self.iter_entries()
            .filter(move |entry| entry.object_type() == object_type)
    }

    /// Collect all entries in deterministic numeric object-ID order.
    #[cfg(test)]
    pub fn all_entries(&self) -> Vec<ObjectIndexEntry<'_>> {
        self.iter_entries().collect()
    }

    /// Find objects by type in deterministic numeric object-ID order.
    #[cfg(test)]
    pub fn find_objects_by_type(&self, object_type: u32) -> Vec<ObjectIndexEntry<'_>> {
        self.iter_entries_by_type(object_type).collect()
    }

    /// Get typed dependencies without exposing raw sentinel IDs.
    #[cfg(test)]
    pub fn dependencies(&self, object_id: ObjectId) -> Option<impl Iterator<Item = ObjectId> + '_> {
        self.snapshot.locations.outgoing(object_id)
    }

    /// Get typed dependents without exposing raw sentinel IDs.
    #[cfg(test)]
    pub fn dependents(&self, object_id: ObjectId) -> Option<impl Iterator<Item = ObjectId> + '_> {
        self.snapshot.locations.incoming(object_id)
    }

    /// Borrow every indexed edge as a validated typed reference.
    ///
    /// Graph storage and native identifiers remain private to the adapter and
    /// neutral index; callers receive only the stable semantic edge values.
    #[cfg(test)]
    pub fn references(&self) -> impl Iterator<Item = Reference> + '_ {
        self.snapshot.locations.references()
    }

    /// Check for a cycle through the validated identity API.
    #[cfg(test)]
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
    #[cfg(test)]
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
    /// payloads. The view cannot outlive `bundle`.
    pub fn resolve_ref<'a>(
        &self,
        bundle: &'a Bundle,
        object_id: ObjectId,
    ) -> Result<Option<ResolvedObjectRef<'a>>> {
        let Some(entry) = self.entry(object_id) else {
            return Ok(None);
        };

        let fragment_name = self.fragment_name(entry.fragment_id())?;

        let Some(archive) = bundle.get_archive(fragment_name) else {
            return Err(Error::Bundle(format!("Archive {fragment_name} not found")));
        };

        let object = indexed_object(archive, &entry, object_id, fragment_name)?;

        Ok(Some(ResolvedObjectRef {
            id: object_id,
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

    /// Borrow multiple indexed objects in the caller's request order.
    #[cfg(test)]
    pub fn resolve_many_refs<'a>(
        &self,
        bundle: &'a Bundle,
        object_ids: &[ObjectId],
    ) -> Result<Vec<ResolvedObjectRef<'a>>> {
        let mut requested = HashSet::new();
        reserve_set(
            &mut requested,
            object_ids.len(),
            "object index batch request catalog",
        )?;
        let mut requests = Vec::new();
        reserve_vec_exact(
            &mut requests,
            object_ids.len(),
            "object index batch request storage",
        )?;
        for (request_position, object_id) in object_ids.iter().enumerate() {
            if !requested.insert(*object_id) {
                return Err(Error::Archive(format!(
                    "object {object_id:?} occurs more than once in a batch"
                )));
            }
            let entry = self.entry(*object_id).ok_or_else(|| {
                Error::Archive(format!(
                    "object {} is not present in the object index",
                    object_id.get()
                ))
            })?;
            requests.push(BatchObject {
                request_position,
                entry,
            });
        }

        let resolved = self.resolve_many_refs_inner(bundle, &requests)?;
        let mut objects = Vec::new();
        reserve_vec_exact(
            &mut objects,
            object_ids.len(),
            "object index batch result storage",
        )?;
        for (request_position, object_id) in object_ids.iter().enumerate() {
            let Some(object) = resolved[request_position] else {
                return Err(Error::Bundle(format!(
                    "object {} could not be resolved from the bundle",
                    object_id.get()
                )));
            };
            objects.push(object);
        }
        Ok(objects)
    }

    /// Resolve grouped batch requests while preserving caller order.
    #[cfg(test)]
    fn resolve_many_refs_inner<'a>(
        &self,
        bundle: &'a Bundle,
        requests: &[BatchObject<'_>],
    ) -> Result<Vec<Option<ResolvedObjectRef<'a>>>> {
        // Group typed requests by fragment to minimize archive lookups. The
        // request position is part of the sort key, retaining caller order
        // within each archive while making fragment traversal deterministic.
        let mut grouped = Vec::new();
        reserve_vec_exact(
            &mut grouped,
            requests.len(),
            "object index grouped batch requests",
        )?;
        grouped.extend_from_slice(requests);
        grouped.sort_unstable_by_key(|request| {
            (request.entry.fragment_id(), request.request_position)
        });

        let mut resolved_slots = Vec::new();
        reserve_vec_exact(
            &mut resolved_slots,
            requests.len(),
            "object index batch resolution slots",
        )?;
        resolved_slots.resize(requests.len(), None);

        // Resolve objects archive by archive. The indexed source position
        // avoids rescanning each archive for sparse batches.
        let mut group_start = 0;
        while let Some(first) = grouped.get(group_start) {
            let fragment_id = first.entry.fragment_id();
            let mut group_end = group_start.saturating_add(1);
            while grouped
                .get(group_end)
                .is_some_and(|request| request.entry.fragment_id() == fragment_id)
            {
                group_end = group_end.saturating_add(1);
            }
            let fragment_name = self.fragment_name(fragment_id)?;
            if let Some(archive) = bundle.get_archive(fragment_name) {
                for request in &grouped[group_start..group_end] {
                    let object_id = request.entry.id();
                    let object = indexed_object(archive, &request.entry, object_id, fragment_name)?;
                    let resolved_object = ResolvedObjectRef {
                        id: object_id,
                        messages: &object.messages,
                    };
                    if resolved_slots[request.request_position]
                        .replace(resolved_object)
                        .is_some()
                    {
                        return Err(Error::Archive(format!(
                            "object {} occurs in more than one archive",
                            object_id.get()
                        )));
                    }
                }
            }
            group_start = group_end;
        }

        Ok(resolved_slots)
    }

    /// Check for an indexed object through the validated identity API.
    #[cfg(test)]
    pub fn contains(&self, object_id: ObjectId) -> bool {
        self.snapshot.locations.object(object_id).is_some()
    }

    /// Get the total number of indexed objects
    pub fn object_count(&self) -> usize {
        self.snapshot.locations.len()
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct IndexCardinality {
    fragments: usize,
    objects: usize,
}

/// Count package-scale index inputs before publishing any storage.
///
/// Archive ingress already validates each individual archive, but the index
/// owns aggregate vectors and catalogs. Counting those inputs first lets the
/// adapter reject arithmetic overflow and reserve the complete package shape
/// before traversal starts, rather than growing collections opportunistically
/// as an archive is consumed.
fn preflight_bundle(bundle: &Bundle) -> Result<IndexCardinality> {
    let mut cardinality = IndexCardinality::default();
    for (_archive_name, archive) in bundle.iter_archives() {
        cardinality.fragments = cardinality.fragments.checked_add(1).ok_or_else(|| {
            Error::Archive("object index fragment cardinality overflows usize".to_owned())
        })?;
        cardinality.objects = cardinality
            .objects
            .checked_add(archive.objects.len())
            .ok_or_else(|| {
                Error::Archive("object index object cardinality overflows usize".to_owned())
            })?;
    }
    if cardinality.fragments > u32::MAX as usize {
        return Err(Error::Archive(
            "IWA fragment catalog exceeds u32 capacity".to_owned(),
        ));
    }
    Ok(cardinality)
}

fn reserve_vec_exact<T>(values: &mut Vec<T>, capacity: usize, resource: &str) -> Result<()> {
    values.try_reserve_exact(capacity).map_err(|_error| {
        Error::Archive(format!("could not reserve {resource} for {capacity} items"))
    })
}

fn reserve_map<K, V>(values: &mut HashMap<K, V>, additional: usize, resource: &str) -> Result<()>
where
    K: Eq + std::hash::Hash,
{
    let requested = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::Archive(format!("{resource} cardinality overflows usize")))?;
    values.try_reserve(additional).map_err(|_error| {
        Error::Archive(format!(
            "could not reserve {resource} for {requested} items"
        ))
    })
}

#[cfg(test)]
fn reserve_set<K>(values: &mut HashSet<K>, additional: usize, resource: &str) -> Result<()>
where
    K: Eq + std::hash::Hash,
{
    let requested = values
        .len()
        .checked_add(additional)
        .ok_or_else(|| Error::Archive(format!("{resource} cardinality overflows usize")))?;
    values.try_reserve(additional).map_err(|_error| {
        Error::Archive(format!(
            "could not reserve {resource} for {requested} items"
        ))
    })
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
    builder: &mut IndexBuilder,
    metadata: &mut Vec<PendingObjectMetadata>,
    fragment_names: &HashMap<FragmentId, Arc<str>>,
    object_fragments: &mut HashMap<ObjectId, FragmentId>,
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
        if let Some(existing_fragment_id) = object_fragments.get(&object_id)
            && let Some(existing_name) = fragment_names.get(existing_fragment_id)
        {
            return Err(Error::Archive(format!(
                "object {identifier} occurs in archives {} and {archive_name}",
                existing_name
            )));
        }
        if let Err(error) = builder.add_object(ObjectRecord::new(object_id, fragment_id, span)) {
            if matches!(error, IndexError::DuplicateObject(_))
                && let Some(existing_fragment_id) = object_fragments.get(&object_id)
                && let Some(existing_name) = fragment_names.get(existing_fragment_id)
            {
                return Err(Error::Archive(format!(
                    "object {identifier} occurs in archives {} and {archive_name}",
                    existing_name
                )));
            }
            return Err(index_error(error));
        }
        object_fragments.insert(object_id, fragment_id);
        metadata.push(PendingObjectMetadata {
            id: object_id,
            source_position: ArchiveObjectPosition::new(object_position),
            object_type,
        });

        // MessageInfo is the authoritative, application-independent
        // reference index emitted by iWork for every payload.
        let mut has_indexed_references = false;
        for message_info in &object.archive_info.message_infos {
            has_indexed_references |= !message_info.object_references.is_empty();
            for &reference in &message_info.object_references {
                if let Some(target_id) = ObjectId::new(reference) {
                    add_reference_if_absent(builder, object_id, target_id)?;
                }
            }
        }

        // Some old archives omit MessageInfo references. Decode only
        // unambiguous high-numbered payloads as a compatibility fallback;
        // low message types overlap between Numbers and Keynote.
        if !has_indexed_references && object_type >= 2000 {
            reference_extraction::extract(object_id, object, builder)?;
        }
    }
    Ok(())
}

fn add_reference_if_absent(
    builder: &mut IndexBuilder,
    source: ObjectId,
    target: ObjectId,
) -> Result<bool> {
    builder
        .add_reference_if_absent(source, target)
        .map_err(index_error)
}

fn finish_snapshot(
    builder: IndexBuilder,
    mut metadata: Vec<PendingObjectMetadata>,
    mut fragments: Vec<FragmentIndexEntry>,
) -> Result<Arc<IndexSnapshot>> {
    metadata.sort_unstable_by_key(|entry| entry.id);
    fragments.sort_unstable_by(|left, right| left.name.cmp(&right.name));
    let locations = builder.build_allow_missing_targets().map_err(index_error)?;

    if metadata.len() != locations.len() {
        return Err(Error::Archive(format!(
            "object index metadata count {} does not match neutral record count {}",
            metadata.len(),
            locations.len()
        )));
    }

    let mut fragment_names = HashMap::new();
    reserve_map(
        &mut fragment_names,
        fragments.len(),
        "object index fragment-name catalog",
    )?;
    for fragment in &fragments {
        if fragment_names
            .insert(fragment.id, Arc::clone(&fragment.name))
            .is_some()
        {
            return Err(Error::Archive(format!(
                "object index contains duplicate fragment identity {:?}",
                fragment.id
            )));
        }
    }

    let mut adapter_metadata = Vec::new();
    adapter_metadata
        .try_reserve_exact(metadata.len())
        .map_err(|_| {
            Error::Archive(format!(
                "could not reserve object index adapter metadata for {} objects",
                metadata.len()
            ))
        })?;
    adapter_metadata.extend(metadata.into_iter().map(|entry| ArchiveObjectMetadata {
        source_position: entry.source_position,
        object_type: entry.object_type,
    }));

    Ok(Arc::new(IndexSnapshot {
        locations: Arc::new(locations),
        metadata: Arc::from(adapter_metadata.into_boxed_slice()),
        fragments: Arc::from(fragments.into_boxed_slice()),
        fragment_names: Arc::new(fragment_names),
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
    entry: &ObjectIndexEntry<'_>,
    object_id: ObjectId,
    fragment_name: &str,
) -> Result<&'a ArchiveObject> {
    let source_position = entry.source_position().get();
    let object = archive.objects.get(source_position).ok_or_else(|| {
        Error::Archive(format!(
            "object {} in archive {} has stale source position {}",
            object_id.get(),
            fragment_name,
            source_position
        ))
    })?;

    if object.archive_info.identifier != Some(object_id.get()) {
        return Err(Error::Archive(format!(
            "object {} in archive {} has stale source position {} (found identifier {:?})",
            object_id.get(),
            fragment_name,
            source_position,
            object.archive_info.identifier
        )));
    }

    let observed_span = ByteSpan::new(object.data_offset, object.data_length).map_err(|error| {
        Error::Archive(format!(
            "object {} in archive {} has invalid source span at position {}: {error}",
            object_id.get(),
            fragment_name,
            source_position
        ))
    })?;
    if observed_span != entry.span() {
        return Err(Error::Archive(format!(
            "object {} in archive {} has stale source span at position {} (expected {:?}, found {:?})",
            object_id.get(),
            fragment_name,
            source_position,
            entry.span(),
            observed_span
        )));
    }

    Ok(object)
}

/// A borrowed view of an indexed object and its immutable payloads.
///
/// The view is tied to the private bundle used for resolution. It is the
/// allocation-free read path for traversal and extraction; callers that need
/// an owned value can consume it with [`Self::into_owned`].
#[derive(Debug, Clone, Copy)]
pub struct ResolvedObjectRef<'a> {
    /// Validated object identifier.
    id: ObjectId,
    /// Borrowed raw message data.
    pub messages: &'a [RawMessage],
}

impl ResolvedObjectRef<'_> {
    /// Return the validated object identity.
    pub const fn id(&self) -> ObjectId {
        self.id
    }

    /// Get the primary message type without allocating.
    #[cfg(test)]
    pub fn primary_message_type(&self) -> Option<u32> {
        self.messages.first().map(|message| message.type_)
    }

    /// Iterate over message types without cloning the message payloads.
    pub fn message_types(&self) -> impl Iterator<Item = u32> + '_ {
        self.messages.iter().map(|message| message.type_)
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
        assert!(index.snapshot.metadata.is_empty());
        assert!(index.snapshot.fragments.is_empty());
        assert!(index.snapshot.fragment_names.is_empty());
    }

    #[test]
    fn package_index_preflight_counts_cardinality_before_building() {
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
        let mut package = crate::IWorkPackage::new();
        package
            .replace_archive(
                "Index/Z.iwa",
                &Archive {
                    objects: vec![object(30), object(10)],
                },
            )
            .unwrap();
        package
            .replace_archive(
                "Index/A.iwa",
                &Archive {
                    objects: vec![object(20)],
                },
            )
            .unwrap();
        let bundle = Bundle::from_bytes(&package.to_bytes().unwrap()).unwrap();

        assert_eq!(
            preflight_bundle(&bundle).unwrap(),
            IndexCardinality {
                fragments: 2,
                objects: 3,
            }
        );
        let index = ObjectIndex::from_bundle(&bundle).unwrap();
        assert_eq!(
            index
                .iter_object_ids()
                .map(ObjectId::get)
                .collect::<Vec<_>>(),
            vec![10, 20, 30]
        );
        assert_eq!(
            index
                .snapshot
                .fragment_names
                .values()
                .map(Arc::as_ref)
                .collect::<HashSet<_>>()
                .len(),
            2
        );
    }

    #[test]
    fn package_index_reservations_report_typed_archive_errors() {
        let mut vector = Vec::<u8>::new();
        let vector_error = reserve_vec_exact(&mut vector, usize::MAX, "test vector").unwrap_err();
        assert!(matches!(
            vector_error,
            Error::Archive(message) if message.contains("could not reserve test vector")
        ));

        let mut map = HashMap::<u64, u64>::new();
        let map_error = reserve_map(&mut map, usize::MAX, "test map").unwrap_err();
        assert!(matches!(
            map_error,
            Error::Archive(message) if message.contains("could not reserve test map")
        ));

        let mut set = HashSet::<u64>::new();
        let set_error = reserve_set(&mut set, usize::MAX, "test set").unwrap_err();
        assert!(matches!(
            set_error,
            Error::Archive(message) if message.contains("could not reserve test set")
        ));
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
        let object = ArchiveObject::new(
            123,
            vec![RawMessage {
                type_: 42,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let mut index = ObjectIndex::new();
        index
            .parse_archive(
                "Document.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        let entry = index.entry(ObjectId::try_from(123).unwrap()).unwrap();
        assert_eq!(entry.id().get(), 123);
        assert_eq!(entry.fragment_id(), FragmentId::try_from(1).unwrap());
        assert_eq!(entry.span(), ByteSpan::new(0, 0).unwrap());
        assert_eq!(entry.object_type(), 42);
    }

    #[test]
    fn borrowed_entries_align_adapter_metadata_with_neutral_records() {
        let object = |identifier| {
            ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: u32::try_from(identifier).unwrap(),
                    data: Vec::new(),
                }],
            )
            .unwrap()
        };
        let archive = Archive {
            objects: vec![object(30), object(10)],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Document.iwa", &archive).unwrap();

        let entries = index
            .iter_entries()
            .map(|entry| {
                let object = indexed_object(&archive, &entry, entry.id(), "Document.iwa").unwrap();
                (
                    entry.id().get(),
                    entry.object_type(),
                    object.archive_info.identifier,
                )
            })
            .collect::<Vec<_>>();

        assert_eq!(entries, [(10, 10, Some(10)), (30, 30, Some(30))]);
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
            indexed_object(&original, &entry, object_id, "Index/Test.iwa")
                .unwrap()
                .archive_info
                .identifier,
            Some(10)
        );

        let mut changed_span = original.clone();
        changed_span.objects[0].data_offset = 1;
        let error = indexed_object(&changed_span, &entry, object_id, "Index/Test.iwa").unwrap_err();
        assert!(matches!(
            error,
            Error::Archive(message) if message.contains("stale source span")
        ));

        let reordered = Archive {
            objects: vec![object(20), object(10)],
        };
        let error = indexed_object(&reordered, &entry, object_id, "Index/Test.iwa").unwrap_err();
        assert!(matches!(
            error,
            Error::Archive(message) if message.contains("stale source position")
        ));

        let truncated = Archive {
            objects: vec![object(10)],
        };
        let object_id = ObjectId::try_from(20).unwrap();
        let entry = index.entry(object_id).unwrap();
        let error = indexed_object(&truncated, &entry, object_id, "Index/Test.iwa").unwrap_err();
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
                let object = indexed_object(&archive, &entry, entry.id(), "Index/Benchmark.iwa")?;
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
    fn test_object_index_with_typed_graph_queries() {
        let index = ObjectIndex::new();
        let object_id = ObjectId::try_from(1).unwrap();

        assert_eq!(index.references().next(), None);
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
        object.archive_info.message_infos[0].object_references = vec![0, 30, 20, 30, 0];
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
        assert_eq!(index.snapshot.locations.reference_count(), 2);
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
        assert_eq!(index.snapshot.locations.reference_count(), 0);
    }

    #[test]
    fn fallback_deduplicates_repeated_payload_references() {
        let repeated = Reference {
            identifier: 20,
            ..Default::default()
        };
        let table_data = TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            entries: Vec::new(),
            segments: vec![repeated, repeated],
            ..Default::default()
        };
        let object = ArchiveObject::new(
            10,
            vec![RawMessage {
                type_: 6005,
                data: table_data.encode_to_vec(),
            }],
        )
        .unwrap();
        let archive = Archive {
            objects: vec![object],
        };
        let mut index = ObjectIndex::new();
        index.parse_archive("Index/Test.iwa", &archive).unwrap();

        assert_eq!(
            index
                .dependencies(ObjectId::try_from(10).unwrap())
                .map(|references| references.collect::<Vec<_>>()),
            Some(vec![ObjectId::try_from(20).unwrap()])
        );
        assert_eq!(index.snapshot.locations.reference_count(), 1);
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

        assert_eq!(index.entry(source).map(|entry| entry.id()), Some(source));
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
    }

    #[test]
    fn batch_resolution_reports_stale_fragments_in_deterministic_order() {
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
        let archive = |first, second| Archive {
            objects: vec![object(first), object(second)],
        };

        let mut source_package = crate::IWorkPackage::new();
        source_package
            .replace_archive("Index/B.iwa", &archive(20, 21))
            .unwrap();
        source_package
            .replace_archive("Index/A.iwa", &archive(10, 11))
            .unwrap();
        let source_bundle = Bundle::from_bytes(&source_package.to_bytes().unwrap()).unwrap();
        let index = ObjectIndex::from_bundle(&source_bundle).unwrap();

        let mut stale_package = crate::IWorkPackage::new();
        stale_package
            .replace_archive("Index/B.iwa", &archive(21, 20))
            .unwrap();
        stale_package
            .replace_archive("Index/A.iwa", &archive(11, 10))
            .unwrap();
        let stale_bundle = Bundle::from_bytes(&stale_package.to_bytes().unwrap()).unwrap();

        let error = index
            .resolve_many_refs(
                &stale_bundle,
                &[
                    ObjectId::try_from(21).unwrap(),
                    ObjectId::try_from(10).unwrap(),
                ],
            )
            .unwrap_err();
        assert!(matches!(
            error,
            Error::Archive(message)
                if message.contains("Index/A.iwa")
                    && message.contains("stale source position")
                    && !message.contains("Index/B.iwa")
        ));
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
            .resolve_many_refs(&empty_bundle, &[object_id])
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

        let error = index
            .resolve_many_refs(&empty_bundle, &[object_id])
            .unwrap_err();
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
