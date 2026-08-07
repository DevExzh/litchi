use std::collections::HashSet;
use std::ops::Range;

use litchi_iwa_graph::{ObjectId, ObjectIdIter, ReferenceGraph, ReferenceGraphSnapshot};

use crate::error::{AllocationKind, IndexError};
use crate::{FragmentId, ObjectRecord, Reference};

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
        if self.reference_catalog.contains(&reference) {
            return Err(IndexError::DuplicateReference(reference));
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
        Ok(())
    }

    /// Finish the builder as a deterministic immutable index.
    ///
    /// # Errors
    ///
    /// Returns [`IndexError::UnknownSource`] or [`IndexError::UnknownTarget`]
    /// when a reference endpoint was not registered.
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
    /// registered. Duplicate and allocation failures are reported while
    /// references are added, before this method is called.
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

        self.fragments.sort_unstable();
        self.objects.sort_unstable_by_key(ObjectRecord::id);
        self.references.sort_unstable();

        let mut fragment_pairs = Vec::with_capacity(self.objects.len());
        for object in &self.objects {
            fragment_pairs.push((object.fragment(), object.id()));
        }
        fragment_pairs.sort_unstable();

        let mut fragment_object_ids = Vec::with_capacity(fragment_pairs.len());
        let mut fragment_entries = Vec::with_capacity(self.fragments.len());
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

        let mut graph = ReferenceGraph::new();
        for reference in self.references {
            graph.add_object_reference(reference.source(), reference.target());
        }

        Ok(ObjectIndex {
            objects: self.objects.into_boxed_slice(),
            fragments: fragment_entries.into_boxed_slice(),
            fragment_object_ids: fragment_object_ids.into_boxed_slice(),
            graph: graph.snapshot(),
        })
    }
}

/// An immutable, deterministic object-location and reference index.
///
/// The index stores sorted boxed slices rather than exposing mutable maps.
/// Lookups are binary searches; iteration is stable across processes and does
/// not depend on hash-map order. The graph is also frozen at build time.
#[derive(Debug, Clone)]
pub struct ObjectIndex {
    objects: Box<[ObjectRecord]>,
    fragments: Box<[FragmentEntry]>,
    fragment_object_ids: Box<[ObjectId]>,
    graph: ReferenceGraphSnapshot,
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

    /// Borrow registered fragments in deterministic ordinal order.
    #[must_use]
    pub fn fragments(&self) -> impl ExactSizeIterator<Item = FragmentId> + '_ {
        self.fragments.iter().map(|fragment| fragment.id)
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

    /// Borrow the immutable graph snapshot used for reference queries.
    #[must_use]
    pub fn reference_graph(&self) -> &ReferenceGraphSnapshot {
        &self.graph
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

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "Tests use fixed non-null identities and bounded spans."
)]
mod tests {
    use std::num::NonZeroU32;

    use super::*;
    use crate::{ByteSpan, ByteSpanError, FragmentIdError, ReferenceError};

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
        let graph = index.reference_graph().clone();
        assert!(graph.has_cycle_from(one));
        assert_eq!(graph.object_ids(), [one, two]);
        assert_eq!(
            index
                .object(one)
                .map(ObjectRecord::span)
                .map(ByteSpan::length),
            Some(2)
        );
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
}
