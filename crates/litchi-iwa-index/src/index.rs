use std::collections::HashSet;
use std::fmt;
use std::ops::Range;

use litchi_iwa_graph::{ObjectId, ObjectIdIter, ReferenceGraph, ReferenceGraphSnapshot};

use crate::error::{AllocationKind, FragmentTraversalError, IndexError};
use crate::{FragmentId, FragmentSummary, FragmentTraversalLimit, ObjectRecord, Reference};

#[derive(Debug, Clone)]
struct FragmentEntry {
    id: FragmentId,
    object_range: Range<usize>,
}

/// Mutable, fallible assembler for an immutable [`ObjectIndex`] snapshot.
///
/// Fragment and object registration may be performed in any order relative to
/// reference insertion. References are checked against the completed object
/// catalog when [`Self::build`] is called, so an adapter can translate native
/// records in one traversal without maintaining a second ordering contract.
#[derive(Debug, Default)]
pub struct IndexBuilder {
    fragments: Vec<FragmentId>,
    fragment_catalog: HashSet<FragmentId>,
    objects: Vec<ObjectRecord>,
    object_catalog: HashSet<ObjectId>,
    references: Vec<Reference>,
    reference_catalog: HashSet<Reference>,
    ordered_reference_sources: HashSet<ObjectId>,
}

impl IndexBuilder {
    /// Construct an empty builder.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Register one fragment identity.
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::DuplicateFragment`] when the identity is already
    /// registered, or [`IndexError::Allocation`] when the builder cannot
    /// reserve its next catalog entry.
    pub fn add_fragment(&mut self, fragment: FragmentId) -> Result<(), IndexError> {
        if self.fragment_catalog.contains(&fragment) {
            return Err(IndexError::DuplicateFragment(fragment));
        }
        self.fragments
            .try_reserve(1)
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::Fragments,
                requested: self.fragments.len().saturating_add(1),
            })?;
        self.fragment_catalog
            .try_reserve(1)
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::FragmentCatalog,
                requested: self.fragment_catalog.len().saturating_add(1),
            })?;
        self.fragment_catalog.insert(fragment);
        self.fragments.push(fragment);
        Ok(())
    }

    /// Register one object location.
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::UnknownFragment`] when the object's fragment was
    /// not registered, [`IndexError::DuplicateObject`] for a repeated object
    /// identity, or [`IndexError::Allocation`] when storage cannot grow.
    pub fn add_object(&mut self, object: ObjectRecord) -> Result<(), IndexError> {
        if !self.fragment_catalog.contains(&object.fragment()) {
            return Err(IndexError::UnknownFragment(object.fragment()));
        }
        if self.object_catalog.contains(&object.id()) {
            return Err(IndexError::DuplicateObject(object.id()));
        }
        self.objects
            .try_reserve(1)
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::Objects,
                requested: self.objects.len().saturating_add(1),
            })?;
        self.object_catalog
            .try_reserve(1)
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::ObjectCatalog,
                requested: self.object_catalog.len().saturating_add(1),
            })?;
        self.object_catalog.insert(object.id());
        self.objects.push(object);
        Ok(())
    }

    /// Register one directed object reference.
    ///
    /// The endpoints may be registered before or after this call. Missing
    /// endpoints are validated by the selected build method: [`Self::build`]
    /// requires both endpoints, while [`Self::build_allow_missing_targets`]
    /// permits an absent target. Duplicate references are rejected
    /// immediately.
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::DuplicateReference`] for a repeated edge or
    /// [`IndexError::Allocation`] when the builder cannot grow its storage.
    pub fn add_reference(&mut self, source: ObjectId, target: ObjectId) -> Result<(), IndexError> {
        let reference = Reference::new(source, target);
        if !self.add_reference_if_absent(source, target)? {
            return Err(IndexError::DuplicateReference(reference));
        }
        Ok(())
    }

    /// Register one directed object reference unless it is already present.
    ///
    /// The endpoint and build-time validation semantics are identical to
    /// [`Self::add_reference`]. A newly inserted edge returns `true`; an
    /// existing edge is left unchanged and returns `false`.
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::Allocation`] when the builder cannot grow its
    /// storage. Duplicate references are successful idempotent operations.
    pub fn add_reference_if_absent(
        &mut self,
        source: ObjectId,
        target: ObjectId,
    ) -> Result<bool, IndexError> {
        let reference = Reference::new(source, target);
        if self.reference_catalog.contains(&reference) {
            return Ok(false);
        }
        self.references
            .try_reserve(1)
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::References,
                requested: self.references.len().saturating_add(1),
            })?;
        self.reference_catalog
            .try_reserve(1)
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::ReferenceCatalog,
                requested: self.reference_catalog.len().saturating_add(1),
            })?;
        self.reference_catalog.insert(reference);
        self.references.push(reference);
        Ok(true)
    }

    /// Opt one source into insertion-order outgoing references.
    ///
    /// References are sorted by source and target when the builder is frozen
    /// by default. Some compatibility adapters need to retain the source
    /// order of repeated fields, while still using the same deduplicating
    /// builder and deterministic ordering for every other source. Calling
    /// this method before [`Self::build`] or
    /// [`Self::build_allow_missing_targets`] opts only `source` into that
    /// compatibility behavior.
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::Allocation`] when the source-order catalog cannot
    /// reserve its next entry.
    pub fn preserve_reference_order(&mut self, source: ObjectId) -> Result<(), IndexError> {
        if self.ordered_reference_sources.contains(&source) {
            return Ok(());
        }
        self.ordered_reference_sources
            .try_reserve(1)
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::ReferenceOrderCatalog,
                requested: self.ordered_reference_sources.len().saturating_add(1),
            })?;
        self.ordered_reference_sources.insert(source);
        Ok(())
    }

    /// Finish the builder as a deterministic immutable index.
    ///
    /// References are sorted by source and target unless a source was opted
    /// into insertion order with [`Self::preserve_reference_order`].
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::UnknownSource`] or [`IndexError::UnknownTarget`]
    /// when a reference endpoint was not registered, or
    /// [`IndexError::Allocation`] when derived immutable snapshot storage
    /// cannot be reserved.
    pub fn build(self) -> Result<ObjectIndex, IndexError> {
        self.build_with_target_validation(true)
    }

    /// Finish the builder while preserving references to unindexed targets.
    ///
    /// This is the explicit adapter path for formats that can publish a
    /// reference before its target is present in the current archive set. The
    /// source must still be an indexed object; only the target may be absent.
    /// Dangling targets remain available to graph queries, while
    /// [`ObjectIndex::object`] correctly returns `None` for their missing
    /// location record.
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::UnknownSource`] when a reference source was not
    /// registered, or [`IndexError::Allocation`] when derived immutable
    /// snapshot storage cannot be reserved. Duplicate references are reported
    /// while references are added, before this method is called.
    pub fn build_allow_missing_targets(self) -> Result<ObjectIndex, IndexError> {
        self.build_with_target_validation(false)
    }

    fn build_with_target_validation(
        mut self,
        require_indexed_targets: bool,
    ) -> Result<ObjectIndex, IndexError> {
        for reference in &self.references {
            if !self.object_catalog.contains(&reference.source()) {
                return Err(IndexError::UnknownSource(reference.source()));
            }
            if require_indexed_targets && !self.object_catalog.contains(&reference.target()) {
                return Err(IndexError::UnknownTarget(reference.target()));
            }
        }

        // Duplicate catalogs have completed their validation role. Releasing
        // them before allocating immutable tables keeps builder-only hash
        // storage out of the publication peak.
        drop(self.fragment_catalog);
        drop(self.object_catalog);
        drop(self.reference_catalog);

        self.fragments.sort_unstable();
        self.objects.sort_unstable_by_key(ObjectRecord::id);
        if self.ordered_reference_sources.is_empty() {
            self.references.sort_unstable();
        } else {
            // A stable source sort keeps insertion order among references
            // belonging to an opted-in source. Non-opted sources are sorted
            // by target within their source group to retain the builder's
            // normal deterministic graph order.
            self.references.sort_by_key(|reference| reference.source());
            let mut group_start = 0;
            while let Some(first) = self.references.get(group_start) {
                let source = first.source();
                let mut group_end = group_start.saturating_add(1);
                while self
                    .references
                    .get(group_end)
                    .is_some_and(|reference| reference.source() == source)
                {
                    group_end = group_end.saturating_add(1);
                }
                if !self.ordered_reference_sources.contains(&source) {
                    self.references[group_start..group_end]
                        .sort_unstable_by_key(|reference| reference.target());
                }
                group_start = group_end;
            }
        }

        let mut fragment_pairs =
            try_snapshot_buffer(self.objects.len(), AllocationKind::FragmentObjectPairs)?;
        for object in &self.objects {
            fragment_pairs.push((object.fragment(), object.id()));
        }
        fragment_pairs.sort_unstable();

        let mut fragment_object_ids =
            try_snapshot_buffer(fragment_pairs.len(), AllocationKind::FragmentObjectIds)?;
        let mut fragment_entries =
            try_snapshot_buffer(self.fragments.len(), AllocationKind::FragmentEntries)?;
        let mut pair_position = 0;
        for &fragment in &self.fragments {
            let start = pair_position;
            while let Some((pair_fragment, object_id)) = fragment_pairs.get(pair_position) {
                if *pair_fragment != fragment {
                    break;
                }
                fragment_object_ids.push(*object_id);
                pair_position += 1;
            }
            fragment_entries.push(FragmentEntry {
                id: fragment,
                object_range: start..pair_position,
            });
        }

        drop(fragment_pairs);
        drop(self.fragments);
        drop(self.ordered_reference_sources);

        // Builder vectors deliberately grow geometrically. Moving records into
        // an exactly reserved buffer when needed makes final immutable
        // compaction fallible instead of relying on `Vec::into_boxed_slice` to
        // shrink an over-capacity builder allocation during publication.
        let snapshot_objects = if self.objects.len() == self.objects.capacity() {
            self.objects
        } else {
            let mut objects =
                try_snapshot_buffer(self.objects.len(), AllocationKind::SnapshotObjects)?;
            objects.extend(self.objects);
            objects
        };

        let reference_count = self.references.len();
        let graph = ReferenceGraph::try_from_edges(
            self.references
                .into_iter()
                .map(|reference| (reference.source(), reference.target())),
        )
        .map_err(|_error| IndexError::Allocation {
            kind: AllocationKind::ReferenceGraph,
            requested: reference_count,
        })?;
        let graph = graph
            .try_snapshot()
            .map_err(|_error| IndexError::Allocation {
                kind: AllocationKind::ReferenceGraph,
                requested: reference_count,
            })?;

        Ok(ObjectIndex {
            objects: snapshot_objects.into_boxed_slice(),
            fragments: fragment_entries.into_boxed_slice(),
            fragment_object_ids: fragment_object_ids.into_boxed_slice(),
            graph,
        })
    }
}

/// An immutable, deterministic object-location and reference index.
///
/// The index stores sorted boxed slices rather than exposing mutable maps.
/// Lookups are binary searches; iteration is stable across processes and does
/// not depend on hash-map order. The graph is also frozen at build time, with
/// insertion order retained only for sources explicitly opted in by the
/// builder.
#[derive(Clone)]
pub struct ObjectIndex {
    objects: Box<[ObjectRecord]>,
    fragments: Box<[FragmentEntry]>,
    fragment_object_ids: Box<[ObjectId]>,
    graph: ReferenceGraphSnapshot,
}

impl fmt::Debug for ObjectIndex {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Keep the graph snapshot an implementation detail even through the
        // otherwise useful public Debug implementation. In particular, do
        // not delegate to `ReferenceGraphSnapshot` here: its derived output
        // includes adjacency state that is intentionally absent from this
        // crate's public API.
        formatter
            .debug_struct("ObjectIndex")
            .field("object_count", &self.len())
            .field("fragment_count", &self.fragment_count())
            .field("reference_count", &self.reference_count())
            .finish()
    }
}

impl Default for ObjectIndex {
    fn default() -> Self {
        Self {
            objects: Box::default(),
            fragments: Box::default(),
            fragment_object_ids: Box::default(),
            graph: ReferenceGraph::new().snapshot(),
        }
    }
}

impl ObjectIndex {
    /// Return the number of indexed objects.
    #[must_use]
    pub fn len(&self) -> usize {
        self.objects.len()
    }

    /// Return whether no objects are indexed.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.objects.is_empty()
    }

    /// Return the number of registered fragments.
    #[must_use]
    pub fn fragment_count(&self) -> usize {
        self.fragments.len()
    }

    /// Borrow objects in deterministic object-identity order.
    #[must_use]
    pub fn objects(&self) -> impl ExactSizeIterator<Item = &ObjectRecord> {
        self.objects.iter()
    }

    /// Find one object without exposing an indexing operation.
    #[must_use]
    pub fn object(&self, id: ObjectId) -> Option<&ObjectRecord> {
        self.objects
            .binary_search_by_key(&id, ObjectRecord::id)
            .ok()
            .and_then(|position| self.objects.get(position))
    }

    /// Find one object together with its stable position in this snapshot.
    ///
    /// The position is an ordinal in the immutable, object-ID-sorted record
    /// slice. It is not a source/archive position and must not be used to
    /// address native payload storage. Adapter-owned sidecar metadata can use
    /// it to borrow the record without copying neutral identity or location
    /// fields into a second owner.
    #[must_use]
    pub fn object_with_position(&self, id: ObjectId) -> Option<(usize, &ObjectRecord)> {
        let position = self
            .objects
            .binary_search_by_key(&id, ObjectRecord::id)
            .ok()?;
        self.objects.get(position).map(|record| (position, record))
    }

    /// Borrow an object by its immutable snapshot position.
    ///
    /// This is crate-private so archive-free semantic adapters can derive
    /// their own opaque handles without publishing the native identity or
    /// making the neutral object slice part of their public contract.
    pub(crate) fn object_at_position(&self, position: usize) -> Option<&ObjectRecord> {
        self.objects.get(position)
    }

    /// Borrow registered fragments in deterministic ordinal order.
    #[must_use]
    pub fn fragments(&self) -> impl ExactSizeIterator<Item = FragmentId> + '_ {
        self.fragments.iter().map(|fragment| fragment.id)
    }

    /// Borrow archive-free metadata for every registered fragment.
    ///
    /// Summaries are returned in deterministic adapter-local ordinal order.
    /// They intentionally contain no archive entry names, native identifiers,
    /// source positions, graph state, or payload bytes.
    #[must_use]
    pub fn fragment_summaries(&self) -> impl ExactSizeIterator<Item = FragmentSummary> + '_ {
        self.fragments
            .iter()
            .map(|fragment| FragmentSummary::new(fragment.id, fragment.object_range.len()))
    }

    /// Borrow archive-free metadata for one registered fragment.
    #[must_use]
    pub fn fragment_summary(&self, fragment: FragmentId) -> Option<FragmentSummary> {
        self.fragments
            .binary_search_by_key(&fragment, |entry| entry.id)
            .ok()
            .and_then(|position| self.fragments.get(position))
            .map(|entry| FragmentSummary::new(entry.id, entry.object_range.len()))
    }

    /// Borrow object identities belonging to one fragment.
    #[must_use]
    pub fn fragment_object_ids(&self, fragment: FragmentId) -> Option<&[ObjectId]> {
        let entry = self
            .fragments
            .binary_search_by_key(&fragment, |entry| entry.id)
            .ok()
            .and_then(|position| self.fragments.get(position))?;
        self.fragment_object_ids.get(entry.object_range.clone())
    }

    /// Borrow the object identities for one fragment under an explicit limit.
    ///
    /// The query fails when the complete fragment contains more records than
    /// the supplied limit; it never returns a silently truncated view. The
    /// returned identities are typed and remain owned by this immutable
    /// snapshot. A missing fragment returns `Ok(None)`.
    ///
    /// # Errors
    ///
    /// Returns [`FragmentTraversalError::LimitExceeded`] when the fragment is
    /// larger than `limit`, or [`FragmentTraversalError::InvalidRange`] when
    /// the immutable fragment catalog contains an invalid object range.
    #[must_use = "the bounded fragment query result must be checked"]
    pub fn fragment_object_ids_bounded(
        &self,
        fragment: FragmentId,
        limit: FragmentTraversalLimit,
    ) -> Result<Option<&[ObjectId]>, FragmentTraversalError> {
        let Some(entry) = self
            .fragments
            .binary_search_by_key(&fragment, |entry| entry.id)
            .ok()
            .and_then(|position| self.fragments.get(position))
        else {
            return Ok(None);
        };

        let range = entry.object_range.clone();
        let available = self.fragment_object_ids.len();
        if range.start > range.end || range.end > available {
            return Err(FragmentTraversalError::InvalidRange {
                fragment,
                start: range.start,
                end: range.end,
                available,
            });
        }

        let observed = range.end - range.start;
        if observed > limit.max_objects() {
            return Err(FragmentTraversalError::LimitExceeded {
                fragment,
                observed,
                maximum: limit.max_objects(),
            });
        }

        Ok(Some(&self.fragment_object_ids[range]))
    }

    /// Traverse one fragment's neutral object records under an explicit
    /// object-count limit.
    ///
    /// Records are yielded in deterministic object-identity order and the
    /// iterator allocates no collection. The complete fragment is checked
    /// before iteration begins, so callers either receive every record or a
    /// typed limit error. Missing fragments return `Ok(None)`.
    ///
    /// # Errors
    ///
    /// Returns [`FragmentTraversalError::LimitExceeded`] when the fragment is
    /// larger than `limit`, or [`FragmentTraversalError::MissingObject`] when
    /// the fragment's immutable object catalog is inconsistent. It returns
    /// [`FragmentTraversalError::InvalidRange`] when the fragment catalog
    /// contains an invalid object range.
    #[must_use = "the bounded fragment query result must be checked"]
    pub fn fragment_objects_bounded(
        &self,
        fragment: FragmentId,
        limit: FragmentTraversalLimit,
    ) -> Result<Option<impl Iterator<Item = &ObjectRecord> + '_>, FragmentTraversalError> {
        let Some(object_ids) = self.fragment_object_ids_bounded(fragment, limit)? else {
            return Ok(None);
        };

        // The fragment/object table is derived from the same object catalog
        // during `IndexBuilder::build`, so a missing record indicates a
        // broken immutable snapshot rather than a normal lookup miss. Check
        // the complete slice before publishing an iterator to preserve the
        // complete-or-refused contract and avoid silently dropping records.
        for &object_id in object_ids {
            if self
                .objects
                .binary_search_by_key(&object_id, ObjectRecord::id)
                .is_err()
            {
                return Err(FragmentTraversalError::MissingObject { fragment });
            }
        }

        Ok(Some(object_ids.iter().map(|object_id| {
            let position = self
                .objects
                .binary_search_by_key(object_id, ObjectRecord::id)
                .unwrap_or_else(|_| {
                    unreachable!(
                        "fragment object catalog was validated before iterator publication"
                    )
                });
            &self.objects[position]
        })))
    }

    /// Borrow the immutable graph snapshot used for low-level reference queries.
    ///
    /// This method is retained for source compatibility with the original
    /// index API. It exposes the graph-identity view and is therefore not part
    /// of the archive-free semantic projection; new semantic callers should
    /// use [`crate::SemanticIndex`] and its opaque handles.
    #[deprecated(
        note = "use ObjectIndex::references, outgoing, incoming, reachable, or has_cycle; the graph snapshot is a low-level compatibility view"
    )]
    #[must_use]
    pub fn reference_graph(&self) -> &ReferenceGraphSnapshot {
        &self.graph
    }

    /// Return the number of indexed directed references.
    ///
    /// The count is read from the immutable graph snapshot without exposing
    /// that graph or its adjacency representation to callers.
    #[must_use]
    pub fn reference_count(&self) -> usize {
        self.graph.edge_count()
    }

    /// Borrow every indexed edge as a typed reference in deterministic order.
    ///
    /// Sources are visited in ascending object-identity order. Their outgoing
    /// order is the order frozen by the builder: targets are ascending for
    /// ordinary sources, while sources opted into
    /// [`IndexBuilder::preserve_reference_order`] retain their insertion
    /// order. The iterator allocates no collection and does not expose the
    /// graph's internal adjacency representation.
    #[must_use = "iterating references is required to inspect the indexed edges"]
    pub fn references(&self) -> impl Iterator<Item = Reference> + '_ {
        self.graph.iter_object_ids().flat_map(|source| {
            self.graph
                .outgoing(source)
                .into_iter()
                .flatten()
                .map(move |target| Reference::new(source, target))
        })
    }

    /// Borrow outgoing references for one object without allocating.
    #[must_use]
    pub fn outgoing(&self, source: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.graph.outgoing(source)
    }

    /// Borrow incoming references for one object without allocating.
    #[must_use]
    pub fn incoming(&self, target: ObjectId) -> Option<ObjectIdIter<'_>> {
        self.graph.incoming(target)
    }

    /// Return all objects reachable from one object, including the start.
    #[must_use]
    pub fn reachable(&self, start: ObjectId) -> Vec<ObjectId> {
        self.graph.reachable(start)
    }

    /// Return whether a cycle is reachable from one object.
    #[must_use]
    pub fn has_cycle(&self, start: ObjectId) -> bool {
        self.graph.has_cycle_from(start)
    }
}

fn try_snapshot_buffer<T>(requested: usize, kind: AllocationKind) -> Result<Vec<T>, IndexError> {
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(requested)
        .map_err(|_error| IndexError::Allocation { kind, requested })?;
    Ok(buffer)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "Tests use fixed non-null identities and bounded spans."
)]
mod tests {
    use std::num::NonZeroU32;

    use super::*;
    use crate::{
        ByteSpan, ByteSpanError, FragmentIdError, FragmentTraversalError, FragmentTraversalLimit,
        ReferenceError, SemanticIndex, SemanticLimits,
    };

    fn fragment(value: u32) -> FragmentId {
        FragmentId::new(NonZeroU32::new(value).expect("test fragment is non-zero"))
    }

    fn object(value: u64) -> ObjectId {
        ObjectId::new(value).expect("test object is non-zero")
    }

    fn span(start: u64, length: u64) -> ByteSpan {
        ByteSpan::new(start, length).expect("test span is in bounds")
    }

    #[test]
    fn builder_freezes_sorted_objects_fragments_and_references() {
        let first = fragment(1);
        let second = fragment(2);
        let one = object(1);
        let two = object(2);
        let three = object(3);
        let four = object(4);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(second).expect("second fragment");
        builder.add_fragment(first).expect("first fragment");
        builder
            .add_object(ObjectRecord::new(three, first, span(30, 3)))
            .expect("third object");
        builder
            .add_object(ObjectRecord::new(one, second, span(10, 1)))
            .expect("first object");
        builder
            .add_object(ObjectRecord::new(two, first, span(20, 2)))
            .expect("second object");
        builder
            .add_object(ObjectRecord::new(four, second, span(40, 4)))
            .expect("fourth object");
        builder.add_reference(three, four).expect("third to fourth");
        builder.add_reference(three, one).expect("third to first");
        builder.add_reference(two, three).expect("second to third");

        let index = builder.build().expect("valid index");
        assert_eq!(
            index.objects().map(ObjectRecord::id).collect::<Vec<_>>(),
            [one, two, three, four]
        );
        assert_eq!(index.fragments().collect::<Vec<_>>(), [first, second]);
        assert_eq!(
            index.fragment_summaries().collect::<Vec<_>>(),
            [
                FragmentSummary::new(first, 2),
                FragmentSummary::new(second, 2),
            ]
        );
        assert_eq!(
            index.fragment_summary(first),
            Some(FragmentSummary::new(first, 2))
        );
        assert_eq!(index.fragment_summary(fragment(3)), None);
        assert_eq!(
            index.object_with_position(one).map(|(position, record)| (
                position,
                record.id(),
                record.span()
            )),
            Some((0, one, span(10, 1)))
        );
        assert_eq!(
            index.object_with_position(four).map(|(position, record)| (
                position,
                record.id(),
                record.span()
            )),
            Some((3, four, span(40, 4)))
        );
        assert_eq!(
            index.fragment_object_ids(first),
            Some([two, three].as_slice())
        );
        assert_eq!(
            index.fragment_object_ids(second),
            Some([one, four].as_slice())
        );
        assert_eq!(
            index.outgoing(two).map(Iterator::collect::<Vec<_>>),
            Some(vec![three])
        );
        assert_eq!(
            index.outgoing(three).map(Iterator::collect::<Vec<_>>),
            Some(vec![one, four])
        );
        assert_eq!(
            index.references().collect::<Vec<_>>(),
            [
                Reference::new(two, three),
                Reference::new(three, one),
                Reference::new(three, four),
            ]
        );
        assert_eq!(
            index.incoming(one).map(Iterator::collect::<Vec<_>>),
            Some(vec![three])
        );
        assert_eq!(index.reachable(two), [two, three, one, four]);
        assert!(!index.has_cycle(two));
    }

    #[test]
    fn duplicate_and_missing_values_are_typed_errors() {
        let fragment = fragment(1);
        let first = object(1);
        let second = object(2);
        let mut builder = IndexBuilder::new();
        assert_eq!(builder.add_fragment(fragment), Ok(()));
        assert_eq!(
            builder.add_fragment(fragment),
            Err(IndexError::DuplicateFragment(fragment))
        );
        assert_eq!(
            builder.add_object(ObjectRecord::new(first, fragment, span(0, 1))),
            Ok(())
        );
        assert_eq!(
            builder.add_object(ObjectRecord::new(first, fragment, span(1, 1))),
            Err(IndexError::DuplicateObject(first))
        );
        assert_eq!(builder.add_reference(first, second), Ok(()));
        assert_eq!(
            builder.add_reference(first, second),
            Err(IndexError::DuplicateReference(Reference::new(
                first, second
            )))
        );
        assert!(matches!(
            builder.build(),
            Err(IndexError::UnknownTarget(target)) if target == second
        ));
    }

    #[test]
    fn idempotent_reference_insertion_preserves_strict_duplicate_rejection() {
        let fragment = fragment(1);
        let source = object(1);
        let target = object(2);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_object(ObjectRecord::new(source, fragment, span(0, 1)))
            .expect("source");
        builder
            .add_object(ObjectRecord::new(target, fragment, span(1, 1)))
            .expect("target");

        assert_eq!(builder.add_reference_if_absent(source, target), Ok(true));
        assert_eq!(builder.add_reference_if_absent(source, target), Ok(false));
        assert_eq!(
            builder.add_reference(source, target),
            Err(IndexError::DuplicateReference(Reference::new(
                source, target
            )))
        );

        let index = builder.build().expect("deduplicated index");
        assert_eq!(
            index.outgoing(source).map(Iterator::collect::<Vec<_>>),
            Some(vec![target])
        );
    }

    #[test]
    fn bulk_graph_build_preserves_high_degree_order_and_deduplication() {
        const DEGREE: u64 = 1_024;
        let fragment_id = fragment(1);
        let star_source = object(1);
        let sink = object(DEGREE.saturating_mul(2).saturating_add(2));
        let star_targets = (2..=DEGREE.saturating_add(1))
            .map(object)
            .collect::<Vec<_>>();
        let incoming_sources = (DEGREE.saturating_add(2)
            ..=DEGREE.saturating_mul(2).saturating_add(1))
            .map(object)
            .collect::<Vec<_>>();

        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment_id).expect("fragment");
        builder
            .add_object(ObjectRecord::new(star_source, fragment_id, span(0, 1)))
            .expect("star source");
        for (position, source) in incoming_sources.iter().copied().enumerate() {
            builder
                .add_object(ObjectRecord::new(
                    source,
                    fragment_id,
                    span(u64::try_from(position + 1).unwrap(), 1),
                ))
                .expect("incoming source");
        }
        for target in star_targets.iter().copied() {
            assert!(
                builder
                    .add_reference_if_absent(star_source, target)
                    .expect("star edge")
            );
        }
        for source in incoming_sources.iter().copied() {
            assert!(
                builder
                    .add_reference_if_absent(source, sink)
                    .expect("incoming edge")
            );
        }
        assert!(
            !builder
                .add_reference_if_absent(star_source, star_targets[0])
                .expect("duplicate edge")
        );

        let index = builder
            .build_allow_missing_targets()
            .expect("high-degree index");

        assert_eq!(
            index.outgoing(star_source).map(Iterator::collect::<Vec<_>>),
            Some(star_targets)
        );
        assert_eq!(
            index.incoming(sink).map(Iterator::collect::<Vec<_>>),
            Some(incoming_sources)
        );
        assert_eq!(
            index.reference_count(),
            usize::try_from(DEGREE * 2).unwrap()
        );
    }

    #[test]
    fn opted_in_reference_sources_retain_insertion_order() {
        let fragment = fragment(1);
        let ordered_source = object(1);
        let sorted_source = object(2);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_object(ObjectRecord::new(ordered_source, fragment, span(0, 1)))
            .expect("ordered source");
        builder
            .add_object(ObjectRecord::new(sorted_source, fragment, span(1, 1)))
            .expect("sorted source");

        builder
            .preserve_reference_order(ordered_source)
            .expect("reference-order source");
        for target in [20, 40, 30] {
            assert_eq!(
                builder.add_reference_if_absent(ordered_source, object(target)),
                Ok(true)
            );
        }
        assert_eq!(
            builder.add_reference_if_absent(ordered_source, object(40)),
            Ok(false)
        );
        for target in [4, 2] {
            builder
                .add_reference(sorted_source, object(target))
                .expect("sorted reference");
        }

        let index = builder
            .build_allow_missing_targets()
            .expect("reference-order index");
        assert_eq!(
            index
                .outgoing(ordered_source)
                .map(Iterator::collect::<Vec<_>>),
            Some(vec![object(20), object(40), object(30)])
        );
        assert_eq!(
            index.references().collect::<Vec<_>>(),
            [
                Reference::new(ordered_source, object(20)),
                Reference::new(ordered_source, object(40)),
                Reference::new(ordered_source, object(30)),
                Reference::new(sorted_source, object(2)),
                Reference::new(sorted_source, object(4)),
            ]
        );
        assert_eq!(
            index
                .outgoing(sorted_source)
                .map(Iterator::collect::<Vec<_>>),
            Some(vec![object(2), object(4)])
        );
    }

    #[test]
    fn builder_rejects_unregistered_fragments_and_reference_sources() {
        let registered_fragment = fragment(1);
        let unregistered_fragment = fragment(2);
        let source = object(1);
        let target = object(2);
        let mut builder = IndexBuilder::new();
        builder
            .add_fragment(registered_fragment)
            .expect("registered fragment");

        assert_eq!(
            builder.add_object(ObjectRecord::new(source, unregistered_fragment, span(0, 1),)),
            Err(IndexError::UnknownFragment(unregistered_fragment))
        );

        builder
            .add_object(ObjectRecord::new(target, registered_fragment, span(0, 1)))
            .expect("target object");
        builder
            .add_reference(source, target)
            .expect("reference with deferred source validation");

        assert!(matches!(
            builder.build(),
            Err(IndexError::UnknownSource(missing)) if missing == source
        ));
    }

    #[test]
    fn null_and_overflow_inputs_are_rejected_before_indexing() {
        assert_eq!(FragmentId::try_from(0), Err(FragmentIdError::Null));
        assert_eq!(
            Reference::try_new(None, ObjectId::new(1)),
            Err(ReferenceError::NullSource)
        );
        assert_eq!(
            Reference::try_new(ObjectId::new(1), None),
            Err(ReferenceError::NullTarget)
        );
        assert_eq!(
            ByteSpan::new(u64::MAX, 1),
            Err(ByteSpanError::Overflow {
                start: u64::MAX,
                length: 1
            })
        );
        assert_eq!(
            ByteSpan::from_endpoints(4, 3),
            Err(ByteSpanError::Reversed { start: 4, end: 3 })
        );
    }

    #[test]
    fn graph_snapshot_is_immutable_and_cycle_queries_are_typed() {
        let fragment = fragment(1);
        let one = object(1);
        let two = object(2);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_object(ObjectRecord::new(one, fragment, span(0, 2)))
            .expect("first object");
        builder
            .add_object(ObjectRecord::new(two, fragment, span(2, 2)))
            .expect("second object");
        builder.add_reference(one, two).expect("first edge");
        builder.add_reference(two, one).expect("second edge");
        let index = builder.build().expect("valid index");
        assert!(index.has_cycle(one));
        assert_eq!(index.reference_count(), 2);
        assert_eq!(index.references().count(), 2);
        assert_eq!(
            index
                .object(one)
                .map(ObjectRecord::span)
                .map(ByteSpan::length),
            Some(2)
        );
    }

    #[test]
    #[allow(
        deprecated,
        reason = "the test exercises the documented compatibility bridge"
    )]
    fn legacy_graph_snapshot_bridge_retains_the_original_view() {
        let fragment = fragment(1);
        let source = object(1);
        let target = object(2);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_object(ObjectRecord::new(source, fragment, span(0, 1)))
            .expect("source");
        builder
            .add_object(ObjectRecord::new(target, fragment, span(1, 1)))
            .expect("target");
        builder.add_reference(source, target).expect("reference");

        let index = builder.build().expect("valid index");
        let snapshot: ReferenceGraphSnapshot = index.reference_graph().clone();
        assert_eq!(snapshot.object_ids(), [source, target]);
        assert_eq!(
            snapshot.outgoing(source).map(Iterator::collect),
            Some(vec![target])
        );
    }

    #[test]
    fn debug_output_does_not_publish_private_graph_state() {
        let fragment = fragment(1);
        let source = object(1);
        let target = object(2);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_object(ObjectRecord::new(source, fragment, span(0, 1)))
            .expect("source");
        builder
            .add_object(ObjectRecord::new(target, fragment, span(1, 1)))
            .expect("target");
        builder.add_reference(source, target).expect("reference");

        let debug = format!("{:?}", builder.build().expect("index"));
        assert_eq!(
            debug,
            "ObjectIndex { object_count: 2, fragment_count: 1, reference_count: 1 }"
        );
        assert!(!debug.contains("ReferenceGraph"));
        assert!(!debug.contains("outgoing_refs"));
    }

    #[test]
    fn dangling_targets_remain_queryable_without_object_records() {
        let fragment = fragment(1);
        let source = object(1);
        let dangling = object(99);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_object(ObjectRecord::new(source, fragment, span(0, 4)))
            .expect("source");
        builder
            .add_reference(source, dangling)
            .expect("dangling edge");

        let index = builder
            .build_allow_missing_targets()
            .expect("dangling targets are supported");

        assert_eq!(index.len(), 1);
        assert_eq!(index.object(source).map(ObjectRecord::id), Some(source));
        assert_eq!(index.object(dangling), None);
        assert_eq!(
            index.outgoing(source).map(Iterator::collect::<Vec<_>>),
            Some(vec![dangling])
        );
        assert_eq!(
            index.incoming(dangling).map(Iterator::collect::<Vec<_>>),
            Some(vec![source])
        );
        assert_eq!(index.reachable(source), [source, dangling]);
    }

    #[test]
    fn dangling_target_mode_still_rejects_unindexed_sources() {
        let fragment = fragment(1);
        let source = object(1);
        let target = object(2);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment).expect("fragment");
        builder
            .add_reference(source, target)
            .expect("reference can precede object registration");

        assert!(matches!(
            builder.build_allow_missing_targets(),
            Err(IndexError::UnknownSource(actual)) if actual == source
        ));
    }

    #[test]
    fn bounded_fragment_traversal_is_complete_or_refused() {
        let first = fragment(1);
        let second = fragment(2);
        let one = object(1);
        let two = object(2);
        let three = object(3);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(second).expect("second fragment");
        builder.add_fragment(first).expect("first fragment");
        builder
            .add_object(ObjectRecord::new(three, first, span(30, 3)))
            .expect("third object");
        builder
            .add_object(ObjectRecord::new(one, first, span(10, 1)))
            .expect("first object");
        builder
            .add_object(ObjectRecord::new(two, second, span(20, 2)))
            .expect("second object");

        let index = builder.build().expect("valid index");
        assert_eq!(
            index.fragment_object_ids_bounded(first, FragmentTraversalLimit::new(1)),
            Err(FragmentTraversalError::LimitExceeded {
                fragment: first,
                observed: 2,
                maximum: 1,
            })
        );
        assert_eq!(
            index.fragment_object_ids_bounded(first, FragmentTraversalLimit::new(0)),
            Err(FragmentTraversalError::LimitExceeded {
                fragment: first,
                observed: 2,
                maximum: 0,
            })
        );
        assert_eq!(
            index
                .fragment_objects_bounded(first, FragmentTraversalLimit::new(2))
                .expect("bounded fragment")
                .expect("registered fragment")
                .map(ObjectRecord::id)
                .collect::<Vec<_>>(),
            [one, three]
        );
        assert_eq!(
            index
                .fragment_objects_bounded(second, FragmentTraversalLimit::new(1))
                .expect("bounded fragment")
                .expect("registered fragment")
                .map(ObjectRecord::id)
                .collect::<Vec<_>>(),
            [two]
        );
        assert!(
            index
                .fragment_objects_bounded(fragment(99), FragmentTraversalLimit::new(0))
                .expect("missing fragments are not errors")
                .is_none()
        );
    }

    #[test]
    fn zero_fragment_limit_accepts_only_empty_fragments() {
        let empty = fragment(1);
        let populated = fragment(2);
        let object_id = object(1);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(populated).expect("populated fragment");
        builder.add_fragment(empty).expect("empty fragment");
        builder
            .add_object(ObjectRecord::new(object_id, populated, span(0, 1)))
            .expect("object");

        let index = builder.build().expect("valid index");
        assert_eq!(
            index
                .fragment_object_ids_bounded(empty, FragmentTraversalLimit::new(0))
                .expect("empty fragment fits"),
            Some([].as_slice())
        );
        assert!(matches!(
            index.fragment_object_ids_bounded(populated, FragmentTraversalLimit::new(0)),
            Err(FragmentTraversalError::LimitExceeded {
                fragment: actual,
                observed: 1,
                maximum: 0,
            }) if actual == populated
        ));
    }

    #[test]
    fn semantic_fragment_traversal_reports_missing_object_invariant() {
        let fragment_id = fragment(1);
        let missing_object = object(9);
        let index = ObjectIndex {
            objects: Box::default(),
            fragments: vec![FragmentEntry {
                id: fragment_id,
                object_range: 0..1,
            }]
            .into_boxed_slice(),
            fragment_object_ids: vec![missing_object].into_boxed_slice(),
            graph: ReferenceGraph::new().snapshot(),
        };
        let view = SemanticIndex::new(&index, SemanticLimits::new(0, 0, 0))
            .expect("empty object and reference catalogs fit the view");

        assert!(matches!(
            view.fragment_objects_bounded(fragment_id, FragmentTraversalLimit::new(1)),
            Err(FragmentTraversalError::MissingObject {
                fragment: actual_fragment,
            }) if actual_fragment == fragment_id
        ));
    }

    #[test]
    fn bounded_fragment_queries_reject_invalid_catalog_ranges() {
        let fragment_id = fragment(1);
        let index = ObjectIndex {
            objects: Box::default(),
            fragments: vec![FragmentEntry {
                id: fragment_id,
                object_range: 1..2,
            }]
            .into_boxed_slice(),
            fragment_object_ids: vec![object(1)].into_boxed_slice(),
            graph: ReferenceGraph::new().snapshot(),
        };
        let expected = FragmentTraversalError::InvalidRange {
            fragment: fragment_id,
            start: 1,
            end: 2,
            available: 1,
        };

        assert_eq!(
            index.fragment_object_ids_bounded(fragment_id, FragmentTraversalLimit::new(1)),
            Err(expected)
        );
        assert!(matches!(
            index.fragment_objects_bounded(fragment_id, FragmentTraversalLimit::new(1)),
            Err(actual) if actual == expected
        ));
    }

    #[test]
    fn snapshot_reservation_failures_identify_each_derived_storage() {
        for kind in [
            AllocationKind::SnapshotObjects,
            AllocationKind::FragmentObjectPairs,
            AllocationKind::FragmentObjectIds,
            AllocationKind::FragmentEntries,
        ] {
            assert_eq!(
                try_snapshot_buffer::<u8>(usize::MAX, kind),
                Err(IndexError::Allocation {
                    kind,
                    requested: usize::MAX,
                })
            );
        }
    }
}
