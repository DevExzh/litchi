//! Archive-free semantic views over an immutable object index.
//!
//! [`ObjectIndex`] is intentionally a neutral physical
//! substrate: its graph identities are useful to format adapters, but they
//! are still close to the native wire boundary. This module provides a
//! smaller hand-off for callers that need only semantic object handles and
//! location metadata. It borrows the existing snapshot, keeps no archive
//! state, and never publishes a native graph identity.

use std::collections::{HashSet, VecDeque};
use std::fmt;
use std::num::NonZeroUsize;
use std::sync::atomic::{AtomicUsize, Ordering};

use crate::{
    ByteSpan, FragmentId, FragmentSummary, FragmentTraversalError, FragmentTraversalLimit,
    ObjectId, ObjectIndex, ObjectRecord,
};

static NEXT_PROVENANCE: AtomicUsize = AtomicUsize::new(1);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[repr(C)]
struct HandleData {
    ordinal: NonZeroUsize,
    provenance: NonZeroUsize,
}

/// An adapter-local handle for one object in one immutable index snapshot.
///
/// Handles are one-based positions in the index's deterministic object order.
/// They are not native object identifiers, archive positions, or byte
/// offsets, and they are valid only with the [`SemanticIndex`] that produced
/// them. Each handle also carries a private provenance stamp for that exact
/// semantic view. The stamp is checked when the handle is passed back to a
/// view and is never exposed as a token or serialized value, so an ordinal
/// copied from another view cannot accidentally address this snapshot. The
/// inner ordinal is exposed only as a checked non-zero value so a caller
/// cannot confuse it with a nullable wire identifier.
#[derive(Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[repr(transparent)]
pub struct SemanticObjectHandle(HandleData);

impl fmt::Debug for SemanticObjectHandle {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("SemanticObjectHandle")
            .field(&self.0.ordinal)
            .finish()
    }
}

impl SemanticObjectHandle {
    fn from_position(position: usize, provenance: NonZeroUsize) -> Option<Self> {
        let ordinal = position.checked_add(1)?;
        NonZeroUsize::new(ordinal).map(|ordinal| {
            Self(HandleData {
                ordinal,
                provenance,
            })
        })
    }

    const fn position(self) -> usize {
        self.0.ordinal.get() - 1
    }

    fn belongs_to(self, provenance: NonZeroUsize) -> bool {
        self.0.provenance == provenance
    }

    /// Return this handle's checked adapter-local ordinal.
    ///
    /// The ordinal is meaningful only with the [`SemanticIndex`] that
    /// produced this handle. It is not a globally unique object identifier,
    /// and reconstructing or persisting it without the owning view does not
    /// preserve handle validity.
    #[must_use]
    pub const fn ordinal(self) -> NonZeroUsize {
        self.0.ordinal
    }
}

/// A strict bound for semantic reference traversal.
///
/// The limit counts edge values returned or visited by one operation. It is
/// deliberately separate from [`SemanticLimits`], which caps the complete
/// snapshot at construction time.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct SemanticTraversalLimit(usize);

impl SemanticTraversalLimit {
    /// Construct a reference traversal limit.
    #[must_use]
    pub const fn new(max_references: usize) -> Self {
        Self(max_references)
    }

    /// Return the maximum number of references the operation may inspect.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.0
    }
}

/// Caller-selected ceilings for one borrowed semantic view.
///
/// Construction checks the immutable source before the view is published.
/// No semantic operation can therefore observe more objects or references
/// than the profile admits. A zero ceiling is valid for an empty projection
/// and causes a non-empty source or traversal to fail closed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SemanticLimits {
    max_objects: usize,
    max_references: usize,
    max_reachable_objects: usize,
}

impl SemanticLimits {
    /// Construct a finite semantic profile.
    #[must_use]
    pub const fn new(
        max_objects: usize,
        max_references: usize,
        max_reachable_objects: usize,
    ) -> Self {
        Self {
            max_objects,
            max_references,
            max_reachable_objects,
        }
    }

    /// Return the maximum number of objects admitted by the view.
    #[must_use]
    pub const fn max_objects(self) -> usize {
        self.max_objects
    }

    /// Return the maximum number of references admitted by the view.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.max_references
    }

    /// Return the maximum number of objects returned by reachability.
    #[must_use]
    pub const fn max_reachable_objects(self) -> usize {
        self.max_reachable_objects
    }
}

/// The kind of temporary storage needed by a bounded semantic operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticAllocationKind {
    /// The adapter-local provenance token for one semantic view.
    Provenance,
    /// The immutable reference projection.
    References,
    /// The outgoing-reference projection.
    OutgoingReferences,
    /// The reachability work queue.
    ReachabilityQueue,
    /// The reachability visited set.
    ReachabilityVisited,
    /// The reachability result.
    ReachabilityResult,
}

impl fmt::Display for SemanticAllocationKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::Provenance => "semantic view provenance",
            Self::References => "semantic reference projection",
            Self::OutgoingReferences => "semantic outgoing-reference projection",
            Self::ReachabilityQueue => "semantic reachability queue",
            Self::ReachabilityVisited => "semantic reachability visited set",
            Self::ReachabilityResult => "semantic reachability result",
        };
        formatter.write_str(name)
    }
}

/// Failure while constructing or querying a bounded semantic view.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticError {
    /// The source index contains more objects than the profile permits.
    ObjectLimitExceeded { observed: usize, maximum: usize },
    /// The source index contains more references than the profile permits.
    ReferenceLimitExceeded { observed: usize, maximum: usize },
    /// A reachability query would publish more objects than permitted.
    ReachabilityLimitExceeded { observed: usize, maximum: usize },
    /// A per-query outgoing traversal would inspect too many references.
    TraversalLimitExceeded { observed: usize, maximum: usize },
    /// A handle belongs to another snapshot or is outside this snapshot.
    InvalidHandle(SemanticObjectHandle),
    /// A temporary semantic collection could not reserve its next item.
    Allocation {
        /// The temporary collection that failed.
        kind: SemanticAllocationKind,
        /// The requested item count at the point of failure.
        requested: usize,
    },
}

impl fmt::Display for SemanticError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ObjectLimitExceeded { observed, maximum } => write!(
                formatter,
                "semantic object count {observed} exceeds limit {maximum}"
            ),
            Self::ReferenceLimitExceeded { observed, maximum } => write!(
                formatter,
                "semantic reference count {observed} exceeds limit {maximum}"
            ),
            Self::ReachabilityLimitExceeded { observed, maximum } => write!(
                formatter,
                "semantic reachability count {observed} exceeds limit {maximum}"
            ),
            Self::TraversalLimitExceeded { observed, maximum } => write!(
                formatter,
                "semantic traversal inspected {observed} references, exceeding limit {maximum}"
            ),
            Self::InvalidHandle(handle) => {
                write!(
                    formatter,
                    "semantic handle {handle:?} is not in this snapshot"
                )
            },
            Self::Allocation { kind, requested } => {
                write!(formatter, "could not reserve {kind} for {requested} items")
            },
        }
    }
}

impl std::error::Error for SemanticError {}

/// The semantic target of one reference.
///
/// A target can be absent from the immutable source set. The unresolved state
/// preserves that fact without leaking the missing native identity.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum SemanticTarget {
    /// The target has a location in this semantic snapshot.
    Indexed(SemanticObjectHandle),
    /// The source edge is present but its target is outside this snapshot.
    Unresolved,
}

impl SemanticTarget {
    /// Return the indexed target, if one is present.
    #[must_use]
    pub const fn indexed(self) -> Option<SemanticObjectHandle> {
        match self {
            Self::Indexed(handle) => Some(handle),
            Self::Unresolved => None,
        }
    }

    /// Return whether this target is outside the source snapshot.
    #[must_use]
    pub const fn is_unresolved(self) -> bool {
        matches!(self, Self::Unresolved)
    }
}

/// One archive-free semantic edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct SemanticReference {
    source: SemanticObjectHandle,
    target: SemanticTarget,
}

impl SemanticReference {
    const fn new(source: SemanticObjectHandle, target: SemanticTarget) -> Self {
        Self { source, target }
    }

    /// Return the source handle.
    #[must_use]
    pub const fn source(self) -> SemanticObjectHandle {
        self.source
    }

    /// Return the target state without exposing a native identity.
    #[must_use]
    pub const fn target(self) -> SemanticTarget {
        self.target
    }
}

/// A borrowed semantic object view.
///
/// Only the adapter-local handle and neutral location metadata are exposed.
/// The underlying [`ObjectRecord`] remains private, so its native graph
/// identity cannot cross this semantic boundary.
#[derive(Clone, Copy)]
pub struct SemanticObject<'a> {
    handle: SemanticObjectHandle,
    record: &'a ObjectRecord,
}

impl fmt::Debug for SemanticObject<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SemanticObject")
            .field("handle", &self.handle)
            .field("fragment", &self.record.fragment())
            .field("span", &self.record.span())
            .finish()
    }
}

impl SemanticObject<'_> {
    /// Return the adapter-local handle.
    #[must_use]
    pub const fn handle(self) -> SemanticObjectHandle {
        self.handle
    }

    /// Return the adapter-local fragment identity.
    #[must_use]
    pub const fn fragment(self) -> FragmentId {
        self.record.fragment()
    }

    /// Return the checked byte span within the adapter-local fragment.
    #[must_use]
    pub const fn span(self) -> ByteSpan {
        self.record.span()
    }
}

/// A bounded semantic view over one immutable neutral index.
///
/// The view borrows the index and allocates only bounded result collections.
/// It owns no archive, payload, source name, or native identifier. Handles
/// from one view must not be reused with another view, even if both views
/// happen to have the same object count.
#[derive(Clone, Copy)]
pub struct SemanticIndex<'a> {
    index: &'a ObjectIndex,
    limits: SemanticLimits,
    provenance: NonZeroUsize,
}

impl fmt::Debug for SemanticIndex<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SemanticIndex")
            .field("index", &self.index)
            .field("limits", &self.limits)
            .finish()
    }
}

impl<'a> SemanticIndex<'a> {
    /// Admit an immutable index under an explicit semantic profile.
    ///
    /// # Errors
    ///
    /// Returns a typed limit error before the view is published when the
    /// source snapshot is larger than the selected profile.
    pub fn new(index: &'a ObjectIndex, limits: SemanticLimits) -> Result<Self, SemanticError> {
        if index.len() > limits.max_objects() {
            return Err(SemanticError::ObjectLimitExceeded {
                observed: index.len(),
                maximum: limits.max_objects(),
            });
        }
        if index.reference_count() > limits.max_references() {
            return Err(SemanticError::ReferenceLimitExceeded {
                observed: index.reference_count(),
                maximum: limits.max_references(),
            });
        }
        let provenance = NEXT_PROVENANCE
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |current| {
                current.checked_add(1)
            })
            .map_err(|_current| SemanticError::Allocation {
                kind: SemanticAllocationKind::Provenance,
                requested: 1,
            })
            .and_then(|current| {
                NonZeroUsize::new(current).ok_or(SemanticError::Allocation {
                    kind: SemanticAllocationKind::Provenance,
                    requested: 1,
                })
            })?;
        Ok(Self {
            index,
            limits,
            provenance,
        })
    }

    /// Return the profile that authorized this view.
    #[must_use]
    pub const fn limits(self) -> SemanticLimits {
        self.limits
    }

    /// Return the number of semantic objects.
    #[must_use]
    pub fn len(self) -> usize {
        self.index.len()
    }

    /// Return whether the view contains no semantic objects.
    #[must_use]
    pub fn is_empty(self) -> bool {
        self.index.is_empty()
    }

    /// Borrow archive-free metadata for every registered fragment.
    ///
    /// Fragment summaries contain only the adapter-local fragment ordinal and
    /// the number of neutral object records assigned to it. The iterator
    /// allocates no collection and never exposes the private object IDs used
    /// by the physical index.
    #[must_use]
    pub fn fragment_summaries(self) -> impl ExactSizeIterator<Item = FragmentSummary> + 'a {
        self.index.fragment_summaries()
    }

    /// Borrow archive-free metadata for one registered fragment.
    #[must_use]
    pub fn fragment_summary(self, fragment: FragmentId) -> Option<FragmentSummary> {
        self.index.fragment_summary(fragment)
    }

    /// Traverse one fragment's semantic objects under an explicit record
    /// limit.
    ///
    /// The query is complete or refused: it checks the physical fragment
    /// cardinality before yielding the first item and never truncates the
    /// result. The returned iterator allocates no collection. Raw object IDs
    /// are used only while resolving each private physical record into an
    /// adapter-local semantic handle.
    ///
    /// # Errors
    ///
    /// Returns [`FragmentTraversalError::LimitExceeded`] when the complete
    /// fragment contains more records than `limit`.
    /// Every admitted record is projected; a source/index mismatch is treated
    /// as an internal invariant failure instead of silently dropping a record.
    #[must_use = "the bounded semantic fragment query result must be checked"]
    pub fn fragment_objects_bounded(
        self,
        fragment: FragmentId,
        limit: FragmentTraversalLimit,
    ) -> Result<Option<impl Iterator<Item = SemanticObject<'a>> + 'a>, FragmentTraversalError> {
        let Some(objects) = self.index.fragment_objects_bounded(fragment, limit)? else {
            return Ok(None);
        };

        Ok(Some(objects.map(move |record| {
            let handle = self.handle_for_record(record);
            SemanticObject { handle, record }
        })))
    }

    /// Borrow semantic objects in deterministic adapter-local order.
    /// The projection preserves the immutable index's object cardinality.
    #[must_use = "iterating semantic objects is required to inspect the view"]
    pub fn objects(self) -> impl Iterator<Item = SemanticObject<'a>> + 'a {
        self.index
            .objects()
            .enumerate()
            .map(move |(position, record)| {
                let handle = self.handle_for_position(position);
                SemanticObject { handle, record }
            })
    }

    /// Borrow one semantic object by its adapter-local handle.
    #[must_use]
    pub fn object(self, handle: SemanticObjectHandle) -> Option<SemanticObject<'a>> {
        if !handle.belongs_to(self.provenance) {
            return None;
        }
        let record = self.index.object_at_position(handle.position())?;
        Some(SemanticObject { handle, record })
    }

    /// Project every edge without exposing native endpoint identities.
    ///
    /// The returned boxed slice is allocated only after the source reference
    /// count has passed the profile check. A dangling target is represented by
    /// [`SemanticTarget::Unresolved`], preserving incomplete source sets while
    /// keeping the missing identity private.
    ///
    /// # Errors
    ///
    /// Returns [`SemanticError::Allocation`] when the bounded projection
    /// cannot reserve its result, or [`SemanticError::ReferenceLimitExceeded`]
    /// if a future source snapshot violates the profile.
    #[must_use = "the bounded semantic projection must be checked"]
    pub fn references(self) -> Result<Box<[SemanticReference]>, SemanticError> {
        let observed = self.index.reference_count();
        if observed > self.limits.max_references() {
            return Err(SemanticError::ReferenceLimitExceeded {
                observed,
                maximum: self.limits.max_references(),
            });
        }
        let mut references = Vec::new();
        references
            .try_reserve_exact(observed)
            .map_err(|_error| SemanticError::Allocation {
                kind: SemanticAllocationKind::References,
                requested: observed,
            })?;
        for reference in self.index.references() {
            let Some(source) = self.handle_for_id(reference.source()) else {
                continue;
            };
            let requested = references
                .len()
                .checked_add(1)
                .ok_or(SemanticError::Allocation {
                    kind: SemanticAllocationKind::References,
                    requested: usize::MAX,
                })?;
            if requested > self.limits.max_references() {
                return Err(SemanticError::ReferenceLimitExceeded {
                    observed: requested,
                    maximum: self.limits.max_references(),
                });
            }
            references
                .try_reserve(1)
                .map_err(|_error| SemanticError::Allocation {
                    kind: SemanticAllocationKind::References,
                    requested,
                })?;
            let target = self
                .handle_for_id(reference.target())
                .map_or(SemanticTarget::Unresolved, SemanticTarget::Indexed);
            references.push(SemanticReference::new(source, target));
        }
        Ok(references.into_boxed_slice())
    }

    /// Project one object's outgoing edges under an explicit edge limit.
    ///
    /// The result is complete or refused; it is never silently truncated.
    /// Unresolved targets remain visible as an opaque state.
    ///
    /// # Errors
    ///
    /// Returns [`SemanticError::InvalidHandle`] for a handle from another
    /// snapshot, [`SemanticError::TraversalLimitExceeded`] when the complete
    /// outgoing list exceeds `limit`, or [`SemanticError::Allocation`] when
    /// the bounded result cannot grow.
    #[must_use = "the bounded outgoing projection must be checked"]
    pub fn outgoing(
        self,
        source: SemanticObjectHandle,
        limit: SemanticTraversalLimit,
    ) -> Result<Box<[SemanticTarget]>, SemanticError> {
        let object = self
            .object(source)
            .ok_or(SemanticError::InvalidHandle(source))?;
        let Some(outgoing) = self.index.outgoing(object.record.id()) else {
            return Ok(Box::default());
        };
        let mut targets = Vec::new();
        for target in outgoing {
            let observed = targets.len().saturating_add(1);
            if observed > limit.max_references() {
                return Err(SemanticError::TraversalLimitExceeded {
                    observed,
                    maximum: limit.max_references(),
                });
            }
            targets
                .try_reserve(1)
                .map_err(|_error| SemanticError::Allocation {
                    kind: SemanticAllocationKind::OutgoingReferences,
                    requested: observed,
                })?;
            let target = self
                .handle_for_id(target)
                .map_or(SemanticTarget::Unresolved, SemanticTarget::Indexed);
            targets.push(target);
        }
        Ok(targets.into_boxed_slice())
    }

    /// Return all indexed objects reachable from `start`, including `start`.
    ///
    /// Edges to objects outside the source snapshot are inspected but do not
    /// create a semantic object. Both queue growth and result cardinality are
    /// bounded by the profile; no source graph is mutated.
    ///
    /// # Errors
    ///
    /// Returns [`SemanticError::InvalidHandle`] for a handle from another
    /// snapshot, [`SemanticError::ReachabilityLimitExceeded`] when the result
    /// would exceed its object ceiling, or [`SemanticError::Allocation`] for a
    /// fallible work-collection reservation failure.
    #[must_use = "the bounded reachability result must be checked"]
    pub fn reachable(
        self,
        start: SemanticObjectHandle,
    ) -> Result<Box<[SemanticObjectHandle]>, SemanticError> {
        self.object(start)
            .ok_or(SemanticError::InvalidHandle(start))?;
        if self.limits.max_reachable_objects() == 0 {
            return Err(SemanticError::ReachabilityLimitExceeded {
                observed: 1,
                maximum: 0,
            });
        }

        let mut queue = VecDeque::new();
        queue
            .try_reserve(1)
            .map_err(|_error| SemanticError::Allocation {
                kind: SemanticAllocationKind::ReachabilityQueue,
                requested: 1,
            })?;
        let mut visited = HashSet::new();
        visited
            .try_reserve(1)
            .map_err(|_error| SemanticError::Allocation {
                kind: SemanticAllocationKind::ReachabilityVisited,
                requested: 1,
            })?;
        visited.insert(start);
        queue.push_back(start);

        let mut result = Vec::new();
        let mut traversed_references = 0_usize;
        while let Some(current) = queue.pop_front() {
            let current_object = self
                .object(current)
                .ok_or(SemanticError::InvalidHandle(current))?;
            let result_count = result.len().saturating_add(1);
            if result_count > self.limits.max_reachable_objects() {
                return Err(SemanticError::ReachabilityLimitExceeded {
                    observed: result_count,
                    maximum: self.limits.max_reachable_objects(),
                });
            }
            result
                .try_reserve(1)
                .map_err(|_error| SemanticError::Allocation {
                    kind: SemanticAllocationKind::ReachabilityResult,
                    requested: result_count,
                })?;
            result.push(current);

            let Some(outgoing) = self.index.outgoing(current_object.record.id()) else {
                continue;
            };
            for target_id in outgoing {
                traversed_references = traversed_references.saturating_add(1);
                if traversed_references > self.limits.max_references() {
                    return Err(SemanticError::TraversalLimitExceeded {
                        observed: traversed_references,
                        maximum: self.limits.max_references(),
                    });
                }
                let Some(target) = self.handle_for_id(target_id) else {
                    continue;
                };
                if visited.contains(&target) {
                    continue;
                }
                let next_count = visited.len().saturating_add(1);
                if next_count > self.limits.max_reachable_objects() {
                    return Err(SemanticError::ReachabilityLimitExceeded {
                        observed: next_count,
                        maximum: self.limits.max_reachable_objects(),
                    });
                }
                visited
                    .try_reserve(1)
                    .map_err(|_error| SemanticError::Allocation {
                        kind: SemanticAllocationKind::ReachabilityVisited,
                        requested: next_count,
                    })?;
                queue
                    .try_reserve(1)
                    .map_err(|_error| SemanticError::Allocation {
                        kind: SemanticAllocationKind::ReachabilityQueue,
                        requested: next_count,
                    })?;
                visited.insert(target);
                queue.push_back(target);
            }
        }
        Ok(result.into_boxed_slice())
    }

    fn handle_for_id(self, id: ObjectId) -> Option<SemanticObjectHandle> {
        self.index
            .object_with_position(id)
            .and_then(|(position, _record)| {
                SemanticObjectHandle::from_position(position, self.provenance)
            })
    }

    fn handle_for_position(self, position: usize) -> SemanticObjectHandle {
        match SemanticObjectHandle::from_position(position, self.provenance) {
            Some(handle) => handle,
            None => unreachable!(
                "an immutable object-index position must fit a semantic handle ordinal"
            ),
        }
    }

    fn handle_for_record(self, record: &ObjectRecord) -> SemanticObjectHandle {
        match self.index.object_with_position(record.id()) {
            Some((position, _indexed_record)) => self.handle_for_position(position),
            None => unreachable!(
                "a record yielded by an immutable fragment projection must remain indexed"
            ),
        }
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
    use crate::{
        ByteSpan, FragmentId, FragmentSummary, FragmentTraversalError, FragmentTraversalLimit,
        IndexBuilder,
    };

    fn fragment(value: u32) -> FragmentId {
        FragmentId::new(NonZeroU32::new(value).expect("non-null fragment"))
    }

    fn object(value: u64) -> ObjectId {
        ObjectId::new(value).expect("non-null object")
    }

    fn span(start: u64) -> ByteSpan {
        ByteSpan::new(start, 1).expect("valid span")
    }

    fn index_with_dangling_target() -> ObjectIndex {
        let fragment_id = fragment(1);
        let first = object(10);
        let second = object(20);
        let dangling = object(99);
        let mut builder = IndexBuilder::new();
        builder.add_fragment(fragment_id).expect("fragment");
        builder
            .add_object(ObjectRecord::new(first, fragment_id, span(0)))
            .expect("first object");
        builder
            .add_object(ObjectRecord::new(second, fragment_id, span(1)))
            .expect("second object");
        builder.add_reference(first, second).expect("indexed edge");
        builder
            .add_reference(first, dangling)
            .expect("dangling edge");
        builder.add_reference(second, first).expect("back edge");
        builder
            .build_allow_missing_targets()
            .expect("valid dangling index")
    }

    fn limits(objects: usize, references: usize, reachable: usize) -> SemanticLimits {
        SemanticLimits::new(objects, references, reachable)
    }

    #[test]
    fn semantic_view_uses_opaque_handles_and_neutral_locations() {
        let index = index_with_dangling_target();
        let view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("bounded view");
        let objects = view.objects().collect::<Vec<_>>();
        assert_eq!(objects.len(), 2);
        assert_eq!(objects[0].handle().ordinal().get(), 1);
        assert_eq!(objects[1].handle().ordinal().get(), 2);
        assert_eq!(objects[0].span(), span(0));
        assert_eq!(objects[1].span(), span(1));
        assert_eq!(
            view.object(objects[0].handle())
                .map(SemanticObject::fragment),
            Some(fragment(1))
        );
        let debug = format!("{:?}", objects[0]);
        assert!(!debug.contains("10"));
    }

    #[test]
    fn semantic_fragment_projection_is_bounded_and_hides_object_ids() {
        let index = index_with_dangling_target();
        let view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("bounded view");
        let fragment_id = fragment(1);

        assert_eq!(
            view.fragment_summaries().collect::<Vec<_>>(),
            [FragmentSummary::new(fragment_id, 2)]
        );
        assert_eq!(
            view.fragment_summary(fragment_id),
            Some(FragmentSummary::new(fragment_id, 2))
        );
        assert_eq!(view.fragment_summary(fragment(99)), None);
        assert!(matches!(
            view.fragment_objects_bounded(fragment_id, FragmentTraversalLimit::new(1)),
            Err(FragmentTraversalError::LimitExceeded {
                fragment: actual,
                observed: 2,
                maximum: 1,
            }) if actual == fragment_id
        ));

        let objects = view
            .fragment_objects_bounded(fragment_id, FragmentTraversalLimit::new(2))
            .expect("exact fragment limit")
            .expect("registered fragment")
            .collect::<Vec<_>>();
        assert_eq!(objects.len(), view.len());
        assert_eq!(view.objects().count(), view.len());
        assert_eq!(
            objects
                .iter()
                .copied()
                .map(SemanticObject::handle)
                .collect::<Vec<_>>(),
            view.objects()
                .map(SemanticObject::handle)
                .collect::<Vec<_>>()
        );
        assert_eq!(objects[0].span(), span(0));
        assert_eq!(objects[1].span(), span(1));
        assert!(!format!("{:?}", objects[0]).contains("10"));

        assert!(
            view.fragment_objects_bounded(fragment(99), FragmentTraversalLimit::new(0))
                .expect("missing fragments are not errors")
                .is_none()
        );
    }

    #[test]
    fn semantic_references_hide_dangling_identity() {
        let index = index_with_dangling_target();
        let view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("bounded view");
        let objects = view.objects().collect::<Vec<_>>();
        let references = view.references().expect("reference projection");
        assert_eq!(references.len(), 3);
        assert_eq!(references[0].source().ordinal().get(), 1);
        assert_eq!(
            references[0].target(),
            SemanticTarget::Indexed(objects[1].handle())
        );
        assert_eq!(references[1].target(), SemanticTarget::Unresolved);
        assert_eq!(references[2].source().ordinal().get(), 2);
    }

    #[test]
    fn outgoing_is_complete_or_refused_by_query_limit() {
        let index = index_with_dangling_target();
        let view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("bounded view");
        let objects = view.objects().collect::<Vec<_>>();
        let first = objects[0].handle();
        let second = objects[1].handle();
        assert_eq!(
            view.outgoing(first, SemanticTraversalLimit::new(1)),
            Err(SemanticError::TraversalLimitExceeded {
                observed: 2,
                maximum: 1,
            })
        );
        assert_eq!(
            view.outgoing(first, SemanticTraversalLimit::new(2))
                .expect("exact outgoing limit")
                .as_ref(),
            [SemanticTarget::Indexed(second), SemanticTarget::Unresolved,]
        );
    }

    #[test]
    fn reachability_is_deterministic_and_ignores_only_unresolved_nodes() {
        let index = index_with_dangling_target();
        let view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("bounded view");
        let first = view.objects().next().expect("first object").handle();
        let second = view.objects().nth(1).expect("second object").handle();
        assert_eq!(
            view.reachable(first).expect("reachable").as_ref(),
            [first, second]
        );
        assert_eq!(
            view.reachable(second).expect("reachable").as_ref(),
            [second, first]
        );
    }

    #[test]
    fn view_admission_and_reachability_limits_fail_before_partial_publication() {
        let index = index_with_dangling_target();
        assert!(matches!(
            SemanticIndex::new(&index, limits(1, 3, 1)),
            Err(SemanticError::ObjectLimitExceeded {
                observed: 2,
                maximum: 1,
            })
        ));
        let view = SemanticIndex::new(&index, limits(2, 3, 1)).expect("bounded view");
        let first = view.objects().next().expect("first object").handle();
        assert_eq!(
            view.reachable(first),
            Err(SemanticError::ReachabilityLimitExceeded {
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn invalid_handles_are_typed_and_source_index_stays_unchanged() {
        let index = index_with_dangling_target();
        let view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("bounded view");
        let invalid = SemanticObjectHandle::from_position(2, view.provenance).expect("handle");
        assert!(view.object(invalid).is_none());
        assert_eq!(
            view.outgoing(invalid, SemanticTraversalLimit::new(1)),
            Err(SemanticError::InvalidHandle(invalid))
        );
        assert_eq!(index.len(), 2);
        assert_eq!(index.reference_count(), 3);
    }

    #[test]
    fn handles_are_rejected_across_distinct_views_and_snapshots() {
        let index = index_with_dangling_target();
        let first_view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("first view");
        let copied_view = first_view;
        let second_view = SemanticIndex::new(&index, limits(2, 3, 2)).expect("second view");
        let cloned_snapshot = index.clone();
        let snapshot_view =
            SemanticIndex::new(&cloned_snapshot, limits(2, 3, 2)).expect("snapshot view");
        let first = first_view.objects().next().expect("first object").handle();
        let second = second_view
            .objects()
            .next()
            .expect("second object")
            .handle();

        assert!(copied_view.object(first).is_some());
        assert!(second_view.object(first).is_none());
        assert!(snapshot_view.object(first).is_none());
        assert_ne!(first, second);
        assert_eq!(
            second_view.outgoing(first, SemanticTraversalLimit::new(2)),
            Err(SemanticError::InvalidHandle(first))
        );
        assert_eq!(
            snapshot_view.reachable(first),
            Err(SemanticError::InvalidHandle(first))
        );
    }

    #[test]
    fn zero_reachable_budget_rejects_non_empty_traversal() {
        let index = index_with_dangling_target();
        let view = SemanticIndex::new(&index, limits(2, 3, 0)).expect("bounded view");
        let first = view.objects().next().expect("first object").handle();
        assert_eq!(
            view.reachable(first),
            Err(SemanticError::ReachabilityLimitExceeded {
                observed: 1,
                maximum: 0,
            })
        );
    }
}
