//! Durable exact-source package patches, deterministic merges, and bounded history.

use std::collections::VecDeque;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use super::model::{Limits, Snapshot, invalid, packages_equal, relationships_equal};
use crate::{Error, Result};

mod metadata;
use metadata::Metadata;

const MAGIC: &[u8; 8] = b"LPTX0003";

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipState {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

impl RelationshipState {
    fn capture_bounded(relationships: &litchi_opc::Relationships) -> Result<Vec<Self>> {
        let mut states = Vec::new();
        states
            .try_reserve_exact(relationships.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation patch relationships",
                source,
            })?;
        states.extend(relationships.iter().map(|relationship| Self {
            id: relationship.r_id().to_owned(),
            kind: relationship.reltype().to_owned(),
            target: relationship.target_ref().to_owned(),
            external: relationship.is_external(),
        }));
        states.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Ok(states)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ResourceState {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<RelationshipState>,
}

/// Freeze the visible references used by admission. A foreign part may choose
/// another backing value on a later accessor call; ownership capture must use
/// these exact references rather than re-reading the part after preflight.
#[derive(Clone, Copy)]
struct BorrowedState<'a> {
    part: &'a dyn Part,
    content_type: &'a str,
    blob: &'a [u8],
    relationships: &'a litchi_opc::Relationships,
}

impl<'a> BorrowedState<'a> {
    fn read(part: &'a dyn Part) -> Self {
        Self {
            part,
            content_type: part.content_type(),
            blob: part.blob(),
            relationships: part.rels(),
        }
    }

    fn matches(self, other: Self) -> bool {
        self.content_type == other.content_type
            && (std::ptr::eq(self.blob, other.blob) || self.blob == other.blob)
            && relationships_equal(self.relationships, other.relationships)
    }
}

impl ResourceState {
    fn capture_bounded(part: BorrowedState<'_>) -> Result<Self> {
        let shared = part.part.blob_arc();
        let visible = part.blob;
        let blob = if shared.len() == visible.len() && std::ptr::eq(shared.as_slice(), visible) {
            shared
        } else {
            let mut bytes = Vec::new();
            bytes
                .try_reserve_exact(visible.len())
                .map_err(|source| Error::Allocation {
                    resource: "opened-presentation patch payload",
                    source,
                })?;
            bytes.extend_from_slice(visible);
            Arc::new(bytes)
        };
        Ok(Self {
            content_type: part.content_type.to_owned(),
            blob,
            relationships: RelationshipState::capture_bounded(part.relationships)?,
        })
    }

    fn to_part(&self, name: PackURI) -> Result<BlobPart> {
        let mut part =
            BlobPart::new_shared(name, self.content_type.clone(), Arc::clone(&self.blob));
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.kind.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                if relationship.external {
                    TargetMode::External
                } else {
                    TargetMode::Internal
                },
            )?;
        }
        Ok(part)
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(crate) struct Delta {
    name: PackURI,
    before: Option<ResourceState>,
    after: Option<ResourceState>,
}

impl Delta {
    fn inverse(&self) -> Self {
        Self {
            name: self.name.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// One physical resource conflict in a three-way merge plan.
#[derive(Clone, PartialEq, Eq)]
pub struct Conflict {
    name: PackURI,
}

impl Conflict {
    /// Physical OPC part whose two candidate states differ; `/` names package-root relationships.
    #[must_use]
    pub const fn resource(&self) -> &PackURI {
        &self.name
    }
}

/// Selection for one unresolved three-way conflict.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Resolution {
    /// Retain the immutable base resource.
    Base,
    /// Select the left candidate resource.
    Left,
    /// Select the right candidate resource.
    Right,
}

#[derive(Debug, Clone)]
struct PlannedConflict {
    conflict: Conflict,
    candidates: ConflictCandidates,
    resolution: Option<Resolution>,
}

type RootStates = (Vec<RelationshipState>, Vec<RelationshipState>);

#[derive(Debug, Clone)]
enum ConflictCandidates {
    Part {
        left: Option<Delta>,
        right: Option<Delta>,
        left_metadata: Metadata,
        right_metadata: Metadata,
    },
    Root {
        left: Option<RootStates>,
        right: Option<RootStates>,
        left_metadata: Metadata,
        right_metadata: Metadata,
    },
}
impl PlannedConflict {
    fn selected_metadata(&self) -> Option<&Metadata> {
        let (left, right) = match &self.candidates {
            ConflictCandidates::Part {
                left_metadata,
                right_metadata,
                ..
            }
            | ConflictCandidates::Root {
                left_metadata,
                right_metadata,
                ..
            } => (left_metadata, right_metadata),
        };
        match self.resolution {
            Some(Resolution::Left) => Some(left),
            Some(Resolution::Right) => Some(right),
            _ => None,
        }
    }
    fn selected_part(&self) -> Option<&Delta> {
        match (&self.candidates, self.resolution) {
            (ConflictCandidates::Part { left, .. }, Some(Resolution::Left)) => left.as_ref(),
            (ConflictCandidates::Part { right, .. }, Some(Resolution::Right)) => right.as_ref(),
            _ => None,
        }
    }
}

/// Immutable, non-publishing plan for a deterministic three-way merge.
#[derive(Debug, Clone)]
pub struct ThreeWayPlan {
    presentation_name: PackURI,
    automatic: Vec<Delta>,
    root: Option<RootStates>,
    left_metadata: Metadata,
    right_metadata: Metadata,
    conflicts: Vec<PlannedConflict>,
    limits: Limits,
}

impl ThreeWayPlan {
    /// Unresolved and resolved physical conflicts in deterministic order.
    #[must_use]
    pub fn conflicts(&self) -> impl ExactSizeIterator<Item = &Conflict> {
        self.conflicts.iter().map(|entry| &entry.conflict)
    }

    /// Number of conflicts without a selected candidate.
    #[must_use]
    pub fn unresolved_count(&self) -> usize {
        self.conflicts
            .iter()
            .filter(|entry| entry.resolution.is_none())
            .count()
    }

    /// Return a new plan with one conflict resolved; this plan is unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when `index` is outside the conflict list.
    pub fn resolve(&self, index: usize, resolution: Resolution) -> Result<Self> {
        if index >= self.conflicts.len() {
            return Err(invalid(
                "opened-presentation merge conflict index is out of bounds",
            ));
        }
        let mut next = self.clone();
        let entry = next
            .conflicts
            .get_mut(index)
            .ok_or_else(|| invalid("opened-presentation merge conflict index is out of bounds"))?;
        entry.resolution = Some(resolution);
        Ok(next)
    }

    /// Materialize the selected merge as a durable patch without publication.
    ///
    /// # Errors
    ///
    /// Returns an error while any conflict remains unresolved.
    pub fn finish(&self) -> Result<Patch> {
        if self.unresolved_count() != 0 {
            return Err(invalid(
                "opened-presentation three-way merge has unresolved conflicts",
            ));
        }
        let mut root = self.root.as_ref();
        for entry in &self.conflicts {
            if let ConflictCandidates::Root { left, right, .. } = &entry.candidates {
                match entry.resolution {
                    Some(Resolution::Left) => {
                        root = left.as_ref();
                    },
                    Some(Resolution::Right) => {
                        root = right.as_ref();
                    },
                    Some(Resolution::Base) => {},
                    None => return Err(invalid("unresolved package-root conflict")),
                }
            }
        }
        let selected = self.automatic.iter().chain(
            self.conflicts
                .iter()
                .filter_map(PlannedConflict::selected_part),
        );
        let mut count = 0usize;
        let mut length = MAGIC.len() + 4 + self.presentation_name.as_str().len() + 1 + 4;
        if let Some((before, after)) = root {
            require_relationship_count(before.len(), self.limits)?;
            require_relationship_count(after.len(), self.limits)?;
            add_len(&mut length, relationships_len(before)?)?;
            add_len(&mut length, relationships_len(after)?)?;
        }
        for delta in selected {
            count = count
                .checked_add(1)
                .ok_or_else(|| invalid("patch part count overflow"))?;
            if count > self.limits.max_parts() {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Parts,
                    observed: count,
                    maximum: self.limits.max_parts(),
                });
            }
            add_len(&mut length, 4 + delta.name.as_str().len())?;
            for state in [&delta.before, &delta.after] {
                if let Some(state) = state {
                    require_relationship_count(state.relationships.len(), self.limits)?;
                }
                add_len(&mut length, optional_state_len(state.as_ref())?)?;
            }
            require_capture_bytes(length, self.limits)?;
        }
        add_len(
            &mut length,
            self.left_metadata
                .joined_len(&self.right_metadata, self.limits)?,
        )?;
        let mut groups = self.left_metadata.joined_count(&self.right_metadata);
        for metadata in self
            .conflicts
            .iter()
            .filter_map(PlannedConflict::selected_metadata)
        {
            // Each selected owner was removed from automatic metadata.
            add_len(&mut length, metadata.encoded_len()?.saturating_sub(4))?;
            groups = groups
                .checked_add(metadata.group_count())
                .ok_or_else(|| invalid("metadata count overflow"))?;
        }
        if groups > self.limits.max_parts().saturating_mul(2).saturating_add(2) {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: groups,
                maximum: self.limits.max_parts(),
            });
        }
        require_capture_bytes(length, self.limits)?;
        let mut metadata = self.left_metadata.join(&self.right_metadata, self.limits)?;
        for selected in self
            .conflicts
            .iter()
            .filter_map(PlannedConflict::selected_metadata)
        {
            metadata = metadata.join(selected, self.limits)?;
        }
        let mut deltas = self.automatic.clone();
        deltas.extend(
            self.conflicts
                .iter()
                .filter_map(PlannedConflict::selected_part)
                .cloned(),
        );
        Patch::new(
            self.presentation_name.clone(),
            deltas,
            root.cloned(),
            self.limits,
            metadata,
        )
    }
}

/// Durable, exact-source patch over a finite set of OPC resources.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    pub(crate) presentation_name: PackURI,
    pub(crate) deltas: Vec<Delta>,
    metadata: Metadata,
    root: Option<(Vec<RelationshipState>, Vec<RelationshipState>)>,
    pub(crate) limits: Limits,
    wire_len: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SinglePartChange {
    Remove,
    Add,
}

impl Patch {
    pub(crate) fn capture(
        source: &OpcPackage,
        target: &OpcPackage,
        presentation_name: PackURI,
        limits: Limits,
    ) -> Result<Self> {
        // Admit the complete durable write set using borrowed visible values.
        // No payload owner or relationship string is retained until every
        // changed resource and the aggregate encoded size have passed.
        let root_changed = !relationships_equal(source.rels(), target.rels());
        require_wire_u32(presentation_name.as_str().len())?;
        let mut length = MAGIC.len() + 4 + 1 + 4;
        add_len(&mut length, presentation_name.as_str().len())?;
        if root_changed {
            add_len(
                &mut length,
                borrowed_relationships_len(source.rels(), limits)?,
            )?;
            add_len(
                &mut length,
                borrowed_relationships_len(target.rels(), limits)?,
            )?;
        }
        require_capture_bytes(length, limits)?;
        let mut planned = Vec::new();
        visit_changed_parts(source, target, |name, before, after| {
            if planned.len() >= limits.max_parts() {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Parts,
                    observed: planned.len().checked_add(1).unwrap_or(usize::MAX),
                    maximum: limits.max_parts(),
                });
            }
            require_wire_u32(planned.len().checked_add(1).unwrap_or(usize::MAX))?;
            require_wire_u32(name.as_str().len())?;
            add_len(&mut length, 4)?;
            add_len(&mut length, name.as_str().len())?;
            add_len(&mut length, borrowed_state_len(before, limits)?)?;
            add_len(&mut length, borrowed_state_len(after, limits)?)?;
            require_capture_bytes(length, limits)?;
            planned.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "opened-presentation borrowed patch plan",
                source,
            })?;
            planned.push((name, before, after));
            Ok(())
        })?;
        let metadata = Metadata::capture(source, target, &planned, limits, length)?;
        let mut deltas = Vec::new();
        deltas
            .try_reserve_exact(planned.len())
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation patch deltas",
                source,
            })?;
        for (name, before, after) in planned {
            deltas.push(Delta {
                name: name.clone(),
                before: before.map(ResourceState::capture_bounded).transpose()?,
                after: after.map(ResourceState::capture_bounded).transpose()?,
            });
        }
        let root = if root_changed {
            Some((
                RelationshipState::capture_bounded(source.rels())?,
                RelationshipState::capture_bounded(target.rels())?,
            ))
        } else {
            None
        };
        Self::new(presentation_name, deltas, root, limits, metadata)
    }

    fn new(
        presentation_name: PackURI,
        mut deltas: Vec<Delta>,
        root: Option<(Vec<RelationshipState>, Vec<RelationshipState>)>,
        limits: Limits,
        metadata: Metadata,
    ) -> Result<Self> {
        deltas.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        if deltas.len() > limits.max_parts() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: deltas.len(),
                maximum: limits.max_parts(),
            });
        }
        if deltas.windows(2).any(|pair| pair[0].name == pair[1].name) {
            return Err(invalid(
                "opened-presentation patch contains overlapping part deltas",
            ));
        }
        if deltas.iter().any(|delta| {
            delta.before == delta.after || delta.before.is_none() && delta.after.is_none()
        }) {
            return Err(invalid("opened-presentation patch contains a no-op delta"));
        }
        if root.as_ref().is_some_and(|(before, after)| before == after) {
            return Err(invalid(
                "opened-presentation patch contains a no-op root relationship delta",
            ));
        }
        // Join and three-way resolution intersect policies, so state that was
        // admitted under another patch's limits needs count admission again.
        if let Some((before, after)) = &root {
            require_relationship_count(before.len(), limits)?;
            require_relationship_count(after.len(), limits)?;
        }
        for delta in &deltas {
            for state in [delta.before.as_ref(), delta.after.as_ref()]
                .into_iter()
                .flatten()
            {
                require_relationship_count(state.relationships.len(), limits)?;
            }
        }
        metadata.require_coverage(&deltas, root.as_ref(), limits)?;
        let patch = Self {
            presentation_name,
            deltas,
            root,
            limits,
            metadata,
            wire_len: 0,
        };
        let wire_len = patch.encoded_len_uncached()?;
        if wire_len > patch.limits.max_patch_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::PatchBytes,
                observed: wire_len,
                maximum: patch.limits.max_patch_bytes(),
            });
        }
        Ok(Self { wire_len, ..patch })
    }

    /// Whether this patch changes no OPC resource.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.deltas.is_empty() && self.root.is_none() && self.metadata.is_empty()
    }

    /// Conservative wire-size charge for retaining this patch, without encoding it.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        self.encoded_len().unwrap_or(usize::MAX)
    }

    /// Number of part/root resources plus standalone metadata owners in the write set.
    #[must_use]
    pub fn resource_count(&self) -> usize {
        self.deltas.len()
            + usize::from(self.root.is_some())
            + self
                .metadata
                .additional_resources(&self.deltas, self.root.is_some())
    }

    /// Physical part names in deterministic order.
    #[must_use]
    pub fn resources(&self) -> impl ExactSizeIterator<Item = &PackURI> {
        self.deltas.iter().map(|delta| &delta.name)
    }

    pub(crate) fn removes_resource(&self, name: &PackURI) -> bool {
        self.deltas
            .iter()
            .any(|delta| &delta.name == name && delta.before.is_some() && delta.after.is_none())
    }

    /// The direction and slide of an exact slide removal or its inverse:
    /// the presentation Part changes, one slide Part leaves or returns, and
    /// every other changed resource is a non-slide Part that leaves or
    /// returns with it (the Parts the slide owns, record 0303).
    pub(crate) fn exact_slide_removal_change(&self) -> Option<(SinglePartChange, &PackURI)> {
        if self.root.is_some() || self.deltas.len() < 2 {
            return None;
        }
        let presentation = self
            .deltas
            .iter()
            .find(|delta| delta.name == self.presentation_name)?;
        let (Some(presentation_before), Some(presentation_after)) =
            (&presentation.before, &presentation.after)
        else {
            return None;
        };
        if presentation_before.content_type
            != litchi_opc::constants::content_type::PML_PRESENTATION_MAIN
            || presentation_after.content_type
                != litchi_opc::constants::content_type::PML_PRESENTATION_MAIN
        {
            return None;
        }
        let mut slides = self.deltas.iter().filter(|delta| {
            delta.name != self.presentation_name
                && [delta.before.as_ref(), delta.after.as_ref()]
                    .into_iter()
                    .flatten()
                    .any(|image| {
                        image.content_type == litchi_opc::constants::content_type::PML_SLIDE
                    })
        });
        let slide = slides.next()?;
        if slides.next().is_some() {
            return None;
        }
        let direction = match (&slide.before, &slide.after) {
            (Some(_), None) => SinglePartChange::Remove,
            (None, Some(_)) => SinglePartChange::Add,
            _ => return None,
        };
        let dependents_follow = self
            .deltas
            .iter()
            .filter(|delta| delta.name != self.presentation_name && delta.name != slide.name)
            .all(|delta| match direction {
                SinglePartChange::Remove => delta.before.is_some() && delta.after.is_none(),
                SinglePartChange::Add => delta.before.is_none() && delta.after.is_some(),
            });
        dependents_follow.then_some((direction, &slide.name))
    }

    pub(crate) fn has_same_changes(&self, other: &Self) -> bool {
        self.presentation_name == other.presentation_name
            && self.deltas == other.deltas
            && self.root == other.root
            && self.metadata == other.metadata
    }

    pub(crate) const fn limits(&self) -> Limits {
        self.limits
    }

    /// Exact inverse, suitable only after this patch has been applied.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            presentation_name: self.presentation_name.clone(),
            deltas: self.deltas.iter().map(Delta::inverse).collect(),
            root: self
                .root
                .as_ref()
                .map(|(before, after)| (after.clone(), before.clone())),
            limits: self.limits,
            metadata: self.metadata.inverse(),
            wire_len: self.wire_len,
        }
    }

    /// Whether two patches select different outcomes for any physical resource.
    #[must_use]
    pub fn conflicts_with(&self, other: &Self) -> bool {
        self.metadata.conflicts(&other.metadata)
            || root_conflicts(self.root.as_ref(), other.root.as_ref())
            || self.deltas.iter().any(|left| {
                other
                    .deltas
                    .iter()
                    .any(|right| left.name == right.name && deltas_conflict(left, right))
            })
    }

    /// Deterministically join compatible patches over the same exact base states.
    ///
    /// Identical overlapping outcomes are coalesced. Divergent outcomes or
    /// different before-states are rejected.
    ///
    /// # Errors
    ///
    /// Returns an error for different roots, conflicts, or exceeded bounds.
    pub fn join(&self, other: &Self) -> Result<Self> {
        require_same_root(self, other)?;
        let limits = intersect_limits(self.limits, other.limits)?;
        preflight_join(self, other, limits)?;
        let mut deltas = self.deltas.clone();
        for candidate in &other.deltas {
            if let Some(existing) = deltas.iter_mut().find(|delta| delta.name == candidate.name) {
                *existing = merge_delta(existing, candidate)?;
            } else {
                deltas.push(candidate.clone());
            }
        }
        let root = merge_root(self.root.as_ref(), other.root.as_ref())?;
        if let Some((before, after)) = &root {
            require_relationship_count(before.len(), limits)?;
            require_relationship_count(after.len(), limits)?;
        }
        for delta in &deltas {
            for state in [&delta.before, &delta.after].into_iter().flatten() {
                require_relationship_count(state.relationships.len(), limits)?;
            }
        }

        Self::new(
            self.presentation_name.clone(),
            deltas,
            root,
            limits,
            self.metadata.join(&other.metadata, limits)?,
        )
    }

    /// Build a non-mutating three-way plan against an immutable base snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when either patch is not rooted in `base`.
    pub fn three_way(base: &Snapshot, left: &Self, right: &Self) -> Result<ThreeWayPlan> {
        require_same_root(left, right)?;
        if left.presentation_name != base.presentation_name {
            return Err(invalid(
                "opened-presentation merge patches target a different base root",
            ));
        }
        validate_before(base.package.as_ref(), left)?;
        validate_before(base.package.as_ref(), right)?;
        let limits = intersect_limits(left.limits, right.limits)?;
        // The immutable plan retains both candidates, so each input must fit
        // the intersected policy before any candidate or metadata is cloned.
        for patch in [left, right] {
            require_capture_bytes(patch.encoded_len()?, limits)?;
            if patch.deltas.len() > limits.max_parts() {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Parts,
                    observed: patch.deltas.len(),
                    maximum: limits.max_parts(),
                });
            }
            if let Some((before, after)) = &patch.root {
                require_relationship_count(before.len(), limits)?;
                require_relationship_count(after.len(), limits)?;
            }
            for delta in &patch.deltas {
                for state in [&delta.before, &delta.after].into_iter().flatten() {
                    require_relationship_count(state.relationships.len(), limits)?;
                }
            }
            patch.metadata.check(limits)?;
        }
        let mut automatic = Vec::new();
        let mut conflicts = Vec::new();
        let mut owners = left.metadata.conflicting_part_owners(&right.metadata)?;
        for delta in &left.deltas {
            if right
                .deltas
                .iter()
                .any(|other| other.name == delta.name && deltas_conflict(delta, other))
                && !owners.contains(&delta.name)
            {
                owners.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "opened-presentation conflict owners",
                    source,
                })?;
                owners.push(delta.name.clone());
            }
        }
        let mut left_metadata = left.metadata.clone();
        let mut right_metadata = right.metadata.clone();
        for owner in &owners {
            let a = left.deltas.iter().find(|delta| &delta.name == owner);
            let b = right.deltas.iter().find(|delta| &delta.name == owner);
            if [a, b].into_iter().flatten().any(|delta| {
                delta.before.as_ref().map(|state| &state.content_type)
                    != delta.after.as_ref().map(|state| &state.content_type)
            }) {
                return Err(invalid(
                    "content-type-coupled part conflict resolution is not yet supported",
                ));
            }
            let (left_other, left_selected) =
                left_metadata.partition_relationships(owner, limits)?;
            let (right_other, right_selected) =
                right_metadata.partition_relationships(owner, limits)?;
            left_metadata = left_other;
            right_metadata = right_other;
            conflicts.push(PlannedConflict {
                conflict: Conflict {
                    name: owner.clone(),
                },
                candidates: ConflictCandidates::Part {
                    left: a.cloned(),
                    right: b.cloned(),
                    left_metadata: left_selected,
                    right_metadata: right_selected,
                },
                resolution: None,
            });
        }
        for delta in &left.deltas {
            if owners.contains(&delta.name) {
                continue;
            }
            match right.deltas.iter().find(|other| other.name == delta.name) {
                None => automatic.push(delta.clone()),
                Some(other) if !deltas_conflict(delta, other) => {
                    if let (Some(base), Some(a), Some(b)) =
                        (&delta.before, &delta.after, &other.after)
                    {
                        merged_relationship_len(
                            &base.relationships,
                            &a.relationships,
                            &b.relationships,
                            limits,
                        )?;
                    }
                    automatic.push(merge_delta(delta, other)?)
                },
                Some(_) => return Err(invalid("unclassified part conflict")),
            }
        }
        for delta in &right.deltas {
            if !owners.contains(&delta.name)
                && !left.deltas.iter().any(|other| other.name == delta.name)
            {
                automatic.push(delta.clone());
            }
        }
        let root_conflict = root_conflicts(left.root.as_ref(), right.root.as_ref())
            || left.metadata.root_conflicts(&right.metadata);
        let (left_metadata, right_metadata, root) = if root_conflict {
            let (left_other, left_root) = left_metadata.partition_root(limits)?;
            let (right_other, right_root) = right_metadata.partition_root(limits)?;
            conflicts.push(PlannedConflict {
                conflict: Conflict {
                    name: PackURI::new("/").map_err(Error::Invalid)?,
                },
                candidates: ConflictCandidates::Root {
                    left: left.root.clone(),
                    right: right.root.clone(),
                    left_metadata: left_root,
                    right_metadata: right_root,
                },
                resolution: None,
            });
            (left_other, right_other, None)
        } else {
            if let (Some((base, a)), Some((_, b))) = (&left.root, &right.root) {
                merged_relationship_len(base, a, b, limits)?;
            }
            (
                left_metadata,
                right_metadata,
                merge_root(left.root.as_ref(), right.root.as_ref())?,
            )
        };
        if left_metadata.conflicts(&right_metadata) {
            return Err(invalid(
                "content-type metadata conflict resolution is not yet supported",
            ));
        }
        automatic.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        conflicts.sort_unstable_by(|left, right| {
            left.conflict
                .name
                .as_str()
                .cmp(right.conflict.name.as_str())
        });
        Ok(ThreeWayPlan {
            presentation_name: left.presentation_name.clone(),
            automatic,
            root,
            left_metadata,
            right_metadata,
            conflicts,
            limits,
        })
    }

    /// Serialize this patch into the stable `LPTX0003` binary format.
    ///
    /// # Errors
    ///
    /// Returns an error if lengths overflow or exceed the configured bound.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let length = self.encoded_len()?;
        let mut output = Vec::new();
        output
            .try_reserve_exact(length)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation durable patch",
                source,
            })?;
        self.encode_into(&mut output)?;
        if output.len() == length {
            return Ok(output);
        }

        // The wire length is cached because normal publication is immutable.
        // Tests in this module also deliberately forge a patch by editing its
        // private metadata, though, and that mutation cannot invalidate the
        // cache.  Recompute only on this exceptional mismatch so ordinary
        // serialization keeps the one-pass, cached fast path.
        let fresh_length = self.encoded_len_uncached()?;
        if fresh_length > self.limits.max_patch_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::PatchBytes,
                observed: fresh_length,
                maximum: self.limits.max_patch_bytes(),
            });
        }
        output.clear();
        output
            .try_reserve_exact(fresh_length)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation durable patch",
                source,
            })?;
        self.encode_into(&mut output)?;
        if output.len() != fresh_length {
            return Err(invalid(
                "opened-presentation durable patch length changed during encoding",
            ));
        }
        Ok(output)
    }

    fn encode_into(&self, output: &mut Vec<u8>) -> Result<()> {
        output.extend_from_slice(MAGIC);
        put_bytes32(output, self.presentation_name.as_str().as_bytes())?;
        put_optional_relationship_delta(output, self.root.as_ref())?;
        put_u32(output, self.deltas.len())?;
        for delta in &self.deltas {
            put_bytes32(output, delta.name.as_str().as_bytes())?;
            put_optional_state(output, delta.before.as_ref())?;
            put_optional_state(output, delta.after.as_ref())?;
        }
        self.metadata.encode(output)?;
        Ok(())
    }

    /// Parse a stable durable patch under caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed or unbounded input.
    pub fn from_bytes(bytes: &[u8], limits: Limits) -> Result<Self> {
        if bytes.len() > limits.max_patch_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::PatchBytes,
                observed: bytes.len(),
                maximum: limits.max_patch_bytes(),
            });
        }
        preflight_wire(bytes, limits)?;
        let mut input = Input::new(bytes);
        if input.take(MAGIC.len())? != MAGIC {
            return Err(invalid(
                "opened-presentation durable patch has an unsupported version",
            ));
        }
        let presentation_name = parse_part_name(input.bytes32()?)?;
        let root = parse_optional_relationship_delta(&mut input, limits)?;
        let count = input.usize32()?;
        if count > limits.max_parts() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Parts,
                observed: count,
                maximum: limits.max_parts(),
            });
        }
        let mut deltas = Vec::new();
        deltas
            .try_reserve_exact(count)
            .map_err(|source| Error::Allocation {
                resource: "opened-presentation decoded deltas",
                source,
            })?;
        for _ in 0..count {
            deltas.push(Delta {
                name: parse_part_name(input.bytes32()?)?,
                before: parse_optional_state(&mut input, limits)?,
                after: parse_optional_state(&mut input, limits)?,
            });
        }
        let metadata = Metadata::parse(&mut input, limits)?;
        if !input.is_empty() {
            return Err(invalid(
                "opened-presentation durable patch has trailing bytes",
            ));
        }
        Self::new(presentation_name, deltas, root, limits, metadata)
    }

    pub(crate) fn encoded_len(&self) -> Result<usize> {
        if self.wire_len != 0 {
            return Ok(self.wire_len);
        }
        self.encoded_len_uncached()
    }

    fn encoded_len_uncached(&self) -> Result<usize> {
        let mut length = MAGIC.len();
        add_len(&mut length, 4)?;
        add_len(&mut length, self.presentation_name.as_str().len())?;
        add_len(&mut length, 1)?;
        if let Some((before, after)) = &self.root {
            add_len(&mut length, relationships_len(before)?)?;
            add_len(&mut length, relationships_len(after)?)?;
        }
        add_len(&mut length, 4)?;
        for delta in &self.deltas {
            add_len(&mut length, 4)?;
            add_len(&mut length, delta.name.as_str().len())?;
            add_len(&mut length, optional_state_len(delta.before.as_ref())?)?;
            add_len(&mut length, optional_state_len(delta.after.as_ref())?)?;
            if length > self.limits.max_patch_bytes() {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::PatchBytes,
                    observed: length,
                    maximum: self.limits.max_patch_bytes(),
                });
            }
        }
        add_len(&mut length, self.metadata.encoded_len()?)?;
        Ok(length)
    }
}

/// Visit differences without collecting the complete package's names or
/// asking a foreign `Part` for a possibly unrelated shared payload owner.
fn visit_changed_parts<'a>(
    source: &'a OpcPackage,
    target: &'a OpcPackage,
    mut visit: impl FnMut(
        &'a PackURI,
        Option<BorrowedState<'a>>,
        Option<BorrowedState<'a>>,
    ) -> Result<()>,
) -> Result<()> {
    for before in source.try_iter_parts() {
        let before = before?;
        let name = before.partname();
        let after = optional_part(target, name)?.map(BorrowedState::read);
        let before = BorrowedState::read(before);
        if after.is_none_or(|after| !before.matches(after)) {
            visit(name, Some(before), after)?;
        }
    }
    for after in target.try_iter_parts() {
        let after = after?;
        let name = after.partname();
        if optional_part(source, name)?.is_none() {
            visit(name, None, Some(BorrowedState::read(after)))?;
        }
    }
    Ok(())
}

fn optional_part<'a>(package: &'a OpcPackage, name: &PackURI) -> Result<Option<&'a dyn Part>> {
    match package.get_part(name) {
        Ok(part) => Ok(Some(part)),
        Err(litchi_opc::OpcError::PartNotFound(_)) => Ok(None),
        Err(error) => Err(error.into()),
    }
}

fn require_capture_bytes(length: usize, limits: Limits) -> Result<()> {
    if length > limits.max_patch_bytes() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::PatchBytes,
            observed: length,
            maximum: limits.max_patch_bytes(),
        });
    }
    Ok(())
}

fn borrowed_state_len(part: Option<BorrowedState<'_>>, limits: Limits) -> Result<usize> {
    let mut length = 1usize;
    if let Some(part) = part {
        add_len(&mut length, 4)?;
        require_wire_u32(part.content_type.len())?;
        add_len(&mut length, part.content_type.len())?;
        add_len(&mut length, 8)?;
        add_len(&mut length, part.blob.len())?;
        add_len(
            &mut length,
            borrowed_relationships_len(part.relationships, limits)?,
        )?;
    }
    Ok(length)
}

fn require_wire_u32(length: usize) -> Result<()> {
    u32::try_from(length)
        .map(|_| ())
        .map_err(|_| invalid("opened-presentation durable length exceeds u32"))
}

fn require_relationship_count(count: usize, limits: Limits) -> Result<()> {
    if count > limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Parts,
            observed: count,
            maximum: limits.max_parts(),
        });
    }
    require_wire_u32(count)
}

fn borrowed_relationships_len(
    relationships: &litchi_opc::Relationships,
    limits: Limits,
) -> Result<usize> {
    require_relationship_count(relationships.len(), limits)?;
    let mut length = 4usize;
    for relationship in relationships.iter() {
        for value in [
            relationship.r_id(),
            relationship.reltype(),
            relationship.target_ref(),
        ] {
            require_wire_u32(value.len())?;
            add_len(&mut length, 4)?;
            add_len(&mut length, value.len())?;
        }
        add_len(&mut length, 1)?;
    }
    Ok(length)
}

fn state_matches(part: Option<&dyn Part>, expected: Option<&ResourceState>) -> bool {
    match (part, expected) {
        (None, None) => true,
        (Some(part), Some(expected)) => {
            part.content_type() == expected.content_type
                && part.blob() == expected.blob.as_slice()
                && relationships_match(part.rels(), &expected.relationships)
        },
        _ => false,
    }
}

fn delta_matches(part: Option<&dyn Part>, delta: &Delta, after: bool, metadata: &Metadata) -> bool {
    if metadata.has_relationships(&delta.name)
        && let (Some(part), Some(before), Some(next)) = (part, &delta.before, &delta.after)
    {
        let expected = if after { next } else { before };
        part.content_type() == expected.content_type
            && part.blob() == expected.blob.as_slice()
            && relationship_changes_match(
                part.rels(),
                &before.relationships,
                &next.relationships,
                after,
            )
    } else {
        state_matches(
            part,
            if after {
                delta.after.as_ref()
            } else {
                delta.before.as_ref()
            },
        )
    }
}
fn relationship_changes_match(
    actual: &litchi_opc::Relationships,
    before: &[RelationshipState],
    after: &[RelationshipState],
    target: bool,
) -> bool {
    let matches = |id: &str, expected: Option<&RelationshipState>| match (actual.get(id), expected)
    {
        (None, None) => true,
        (Some(actual), Some(expected)) => {
            actual.reltype() == expected.kind
                && actual.target_ref() == expected.target
                && actual.is_external() == expected.external
        },
        _ => false,
    };
    before.iter().all(|old| {
        let next = after
            .binary_search_by(|value| value.id.cmp(&old.id))
            .ok()
            .map(|index| &after[index]);
        next == Some(old) || matches(&old.id, if target { next } else { Some(old) })
    }) && after
        .iter()
        .filter(|next| {
            before
                .binary_search_by(|value| value.id.cmp(&next.id))
                .is_err()
        })
        .all(|next| matches(&next.id, if target { Some(next) } else { None }))
}

fn relationships_match(actual: &litchi_opc::Relationships, expected: &[RelationshipState]) -> bool {
    actual.len() == expected.len()
        && expected.iter().all(|expected| {
            actual.get(&expected.id).is_some_and(|actual| {
                actual.reltype() == expected.kind
                    && actual.target_ref() == expected.target
                    && actual.is_external() == expected.external
            })
        })
}

fn optional_state_len(value: Option<&ResourceState>) -> Result<usize> {
    let mut length = 1usize;
    if let Some(value) = value {
        add_len(&mut length, 4)?;
        add_len(&mut length, value.content_type.len())?;
        add_len(&mut length, 8)?;
        add_len(&mut length, value.blob.len())?;
        add_len(&mut length, relationships_len(&value.relationships)?)?;
    }
    Ok(length)
}

fn relationships_len(relationships: &[RelationshipState]) -> Result<usize> {
    let mut length = 4usize;
    for relationship in relationships {
        for value in [&relationship.id, &relationship.kind, &relationship.target] {
            add_len(&mut length, 4)?;
            add_len(&mut length, value.len())?;
        }
        add_len(&mut length, 1)?;
    }
    Ok(length)
}

fn add_len(total: &mut usize, value: usize) -> Result<()> {
    *total = total
        .checked_add(value)
        .ok_or_else(|| invalid("opened-presentation durable patch length overflow"))?;
    Ok(())
}

/// Undo/redo history bounded by aggregate durable bytes and entry count.
#[derive(Debug, Clone)]
pub struct History {
    limits: Limits,
    undo: VecDeque<(Patch, usize)>,
    redo: VecDeque<(Patch, usize)>,
    bytes: usize,
}

impl History {
    /// Construct an empty bounded history.
    #[must_use]
    pub fn new(limits: Limits) -> Self {
        Self {
            limits,
            undo: VecDeque::new(),
            redo: VecDeque::new(),
            bytes: 0,
        }
    }

    /// Number of retained undo entries.
    #[must_use]
    pub fn len(&self) -> usize {
        self.undo.len()
    }

    /// Number of retained redo entries.
    #[must_use]
    pub fn redo_len(&self) -> usize {
        self.redo.len()
    }

    /// Whether no undo entry is retained.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.undo.is_empty()
    }

    /// Aggregate durable bytes retained by both stacks.
    #[must_use]
    pub const fn encoded_bytes(&self) -> usize {
        self.bytes
    }

    /// Retain a changed patch and clear the divergent redo branch.
    ///
    /// # Errors
    ///
    /// Returns an error when one patch alone exceeds the history byte bound.
    pub fn push(&mut self, patch: Patch) -> Result<()> {
        if patch.is_noop() {
            return Ok(());
        }
        self.clear_redo();
        let length = patch.encoded_len()?;
        if length > self.limits.max_history_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::HistoryBytes,
                observed: length,
                maximum: self.limits.max_history_bytes(),
            });
        }
        while self.undo.len() >= self.limits.max_history_entries()
            || self
                .bytes
                .checked_add(length)
                .is_none_or(|total| total > self.limits.max_history_bytes())
        {
            let Some((_evicted, removed)) = self.undo.pop_front() else {
                break;
            };
            self.bytes = self.bytes.saturating_sub(removed);
        }
        self.bytes = self
            .bytes
            .checked_add(length)
            .ok_or_else(|| invalid("opened-presentation history byte count overflow"))?;
        self.undo.push_back((patch, length));
        Ok(())
    }

    /// Move the newest forward patch to redo and return its exact inverse.
    pub fn pop_undo(&mut self) -> Option<Patch> {
        let entry = self.undo.pop_back()?;
        let inverse = entry.0.inverse();
        self.redo.push_back(entry);
        Some(inverse)
    }

    /// Move the newest redo patch back to undo and return the forward patch.
    pub fn pop_redo(&mut self) -> Option<Patch> {
        let entry = self.redo.pop_back()?;
        let patch = entry.0.clone();
        self.undo.push_back(entry);
        Some(patch)
    }

    /// Compatibility alias for [`Self::pop_undo`].
    pub fn pop_inverse(&mut self) -> Option<Patch> {
        self.pop_undo()
    }

    fn clear_redo(&mut self) {
        for (_patch, length) in self.redo.drain(..) {
            self.bytes = self.bytes.saturating_sub(length);
        }
    }
}

#[cfg(test)]
pub(crate) fn apply(
    package: &mut OpcPackage,
    patch: &Patch,
    physical_source_provenance: bool,
) -> Result<Snapshot> {
    apply_with_semantic_limits(package, patch, physical_source_provenance, None)
}

pub(crate) fn apply_with_semantic_limits(
    package: &mut OpcPackage,
    patch: &Patch,
    physical_source_provenance: bool,
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
) -> Result<Snapshot> {
    apply_with_revision(
        package,
        patch,
        None,
        None,
        physical_source_provenance,
        None,
        semantic_limits,
    )
}

pub(crate) fn apply_to_new_with_semantic_limits(
    package: &OpcPackage,
    patch: &Patch,
    physical_source_provenance: bool,
    parent: &super::model::PartDigests,
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
) -> Result<(OpcPackage, Snapshot)> {
    applied_with_revision(
        package,
        patch,
        None,
        None,
        physical_source_provenance,
        Some(parent),
        semantic_limits,
    )
}

pub(crate) fn apply_exact_revision_with_semantic_limits(
    package: &mut OpcPackage,
    patch: &Patch,
    result_revision: [u8; 32],
    physical_source_provenance: bool,
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
) -> Result<Snapshot> {
    apply_with_revision(
        package,
        patch,
        Some(result_revision),
        None,
        physical_source_provenance,
        None,
        semantic_limits,
    )
}

/// Capture `candidate`, reusing an already validated snapshot of identical
/// content when the committing transaction supplied one.
fn capture_candidate(
    candidate: &OpcPackage,
    limits: Limits,
    physical_source_provenance: bool,
    committed: Option<&Snapshot>,
    parent: Option<&super::model::PartDigests>,
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
) -> Result<Snapshot> {
    if let Some(committed) = committed
        && committed.limits == limits
        && committed.physical_source_provenance == physical_source_provenance
        && committed.semantic_limits() == crate::open_options::selected(semantic_limits.as_ref())
        && packages_equal(committed.package.as_ref(), candidate)
    {
        return Ok(committed.rebound_to(candidate));
    }
    // A candidate that drifted outside the write set still shares most
    // payload allocations with the committed package, so the fallback capture
    // consults that memo rather than hashing every payload from cold.
    match committed
        .map(|snapshot| snapshot.part_digests.as_ref())
        .or(parent)
    {
        Some(parent) => super::model::capture_with_parent_digests_and_semantic_limits(
            candidate,
            limits,
            physical_source_provenance,
            parent,
            semantic_limits,
        ),
        None => super::model::capture_with_provenance_and_semantic_limits(
            candidate,
            limits,
            physical_source_provenance,
            semantic_limits,
        ),
    }
}

/// Validate a detached candidate already constructed from a freshly proven
/// patch. This retains the application checks without rebuilding that same
/// candidate and discarding its validated physical package representation.
///
/// The candidate is captured here, after `validate_before`, under the patch's
/// limits, with every validation a capture runs. `parent` is a payload-digest
/// memo whose entries name allocations the candidate may hold — the memo of
/// the capture that built it moments earlier in the same operation (change
/// 0751). It is consulted only to skip re-hashing a payload allocation the
/// candidate provably holds; a miss is an ordinary hash, so the snapshot and
/// its revision are the ones a capture without it returns.
pub(crate) fn validate_candidate(
    source: &OpcPackage,
    candidate: &OpcPackage,
    patch: &Patch,
    result_revision: [u8; 32],
    physical_source_provenance: bool,
    parent: Option<&super::model::PartDigests>,
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
) -> Result<Snapshot> {
    let current_main = crate::parts::PresentationPart::from_package(source)?
        .part()
        .partname()
        .clone();
    if current_main != patch.presentation_name {
        return Err(invalid(
            "opened-presentation patch targets a different presentation root",
        ));
    }
    validate_before(source, patch)?;
    let snapshot = match parent {
        Some(parent) => super::model::capture_with_parent_digests_and_semantic_limits(
            candidate,
            patch.limits,
            physical_source_provenance,
            parent,
            semantic_limits,
        )?,
        None => super::model::capture_with_provenance_and_semantic_limits(
            candidate,
            patch.limits,
            physical_source_provenance,
            semantic_limits,
        )?,
    };
    validate_after(candidate, patch)?;
    if snapshot.revision() != result_revision {
        return Err(invalid(
            "opened-presentation candidate has an unexpected complete-package revision",
        ));
    }
    Ok(snapshot)
}

fn apply_with_revision(
    package: &mut OpcPackage,
    patch: &Patch,
    result_revision: Option<[u8; 32]>,
    committed: Option<&Snapshot>,
    physical_source_provenance: bool,
    parent: Option<&super::model::PartDigests>,
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
) -> Result<Snapshot> {
    let (candidate, snapshot) = applied_with_revision(
        package,
        patch,
        result_revision,
        committed,
        physical_source_provenance,
        parent,
        semantic_limits,
    )?;
    *package = candidate;
    Ok(snapshot)
}

fn applied_with_revision(
    package: &OpcPackage,
    patch: &Patch,
    result_revision: Option<[u8; 32]>,
    committed: Option<&Snapshot>,
    physical_source_provenance: bool,
    parent: Option<&super::model::PartDigests>,
    semantic_limits: Option<crate::open_options::SemanticLimitsRef>,
) -> Result<(OpcPackage, Snapshot)> {
    let current_main = crate::parts::PresentationPart::from_package(package)?
        .part()
        .partname()
        .clone();
    if current_main != patch.presentation_name {
        return Err(invalid(
            "opened-presentation patch targets a different presentation root",
        ));
    }
    validate_before(package, patch)?;
    if patch.is_noop() {
        let snapshot = capture_candidate(
            package,
            patch.limits,
            physical_source_provenance,
            committed,
            parent,
            semantic_limits,
        )?;
        if result_revision.is_some_and(|expected| snapshot.revision() != expected) {
            return Err(invalid(
                "opened-presentation candidate has an unexpected complete-package revision",
            ));
        }
        return Ok((package.clone(), snapshot));
    }
    let metadata = patch.metadata.prepare(package, patch.limits)?;
    let mut candidate = package.clone();
    if let Some((_before, after)) = &patch.root
        && !patch
            .metadata
            .has_relationships(&PackURI::new("/").map_err(Error::Invalid)?)
    {
        replace_relationships(candidate.rels_mut(), after)?;
    }
    for delta in &patch.deltas {
        match (&delta.before, &delta.after) {
            (None, Some(state)) => {
                candidate.try_add_part(Box::new(state.to_part(delta.name.clone())?))?;
            },
            (Some(_before), None) => {
                if !candidate.remove_part(&delta.name) {
                    return Err(invalid(
                        "opened-presentation candidate lost a removed resource",
                    ));
                }
            },
            (Some(_before), Some(state)) => {
                let part = candidate.get_part_mut(&delta.name)?;
                if part.content_type() != state.content_type {
                    part.set_content_type(state.content_type.clone())?;
                }
                part.set_blob_shared(Arc::clone(&state.blob));
                if !patch.metadata.has_relationships(&delta.name) {
                    replace_relationships(part.rels_mut(), &state.relationships)?;
                }
            },
            (None, None) => {
                return Err(invalid(
                    "opened-presentation candidate contains a no-op resource",
                ));
            },
        }
    }
    metadata.install(&mut candidate, patch.limits)?;
    validate_topology_dependencies(&candidate, patch)?;
    let snapshot = capture_candidate(
        &candidate,
        patch.limits,
        physical_source_provenance,
        committed,
        parent,
        semantic_limits,
    )?;
    validate_after(&candidate, patch)?;
    if result_revision.is_some_and(|expected| snapshot.revision() != expected) {
        return Err(invalid(
            "opened-presentation candidate has an unexpected complete-package revision",
        ));
    }
    Ok((candidate, snapshot))
}

fn validate_before(package: &OpcPackage, patch: &Patch) -> Result<()> {
    patch.metadata.validate(package, false, patch.limits)?;
    if let Some((before, _after)) = &patch.root
        && !(if patch
            .metadata
            .has_relationships(&PackURI::new("/").map_err(Error::Invalid)?)
        {
            relationship_changes_match(package.rels(), before, _after, false)
        } else {
            relationships_match(package.rels(), before)
        })
    {
        return Err(stale());
    }
    for delta in &patch.deltas {
        let current = optional_part(package, &delta.name)?;
        if !delta_matches(current, delta, false, &patch.metadata) {
            return Err(stale());
        }
    }
    Ok(())
}

fn validate_after(package: &OpcPackage, patch: &Patch) -> Result<()> {
    patch.metadata.validate(package, true, patch.limits)?;
    if let Some((_before, after)) = &patch.root
        && !(if patch
            .metadata
            .has_relationships(&PackURI::new("/").map_err(Error::Invalid)?)
        {
            relationship_changes_match(package.rels(), _before, after, true)
        } else {
            relationships_match(package.rels(), after)
        })
    {
        return Err(invalid(
            "opened-presentation published root relationships differ from the patch target",
        ));
    }
    for delta in &patch.deltas {
        let current = optional_part(package, &delta.name)?;
        if !delta_matches(current, delta, true, &patch.metadata) {
            return Err(invalid(
                "opened-presentation published resource differs from its patch target",
            ));
        }
    }
    Ok(())
}

// Final private-graph admission covers both directions: deletion cannot orphan
// another owner, and creation cannot activate an unrelated dangling edge.
fn validate_topology_dependencies(package: &OpcPackage, patch: &Patch) -> Result<()> {
    let removals = patch.deltas.iter().filter(|d| d.after.is_none()).count();
    let additions = patch.deltas.iter().filter(|d| d.before.is_none()).count();
    if removals == 0 && additions == 0 {
        return Ok(());
    }
    fn folded(value: &str) -> Result<String> {
        let mut text = String::new();
        text.try_reserve_exact(value.len())
            .map_err(|source| Error::Allocation {
                resource: "part dependency name",
                source,
            })?;
        text.push_str(value);
        text.make_ascii_lowercase();
        Ok(text)
    }
    let mut removed = std::collections::HashSet::new();
    let mut added = std::collections::HashSet::new();
    removed
        .try_reserve(removals)
        .map_err(|source| Error::Allocation {
            resource: "removed part dependencies",
            source,
        })?;
    added
        .try_reserve(additions)
        .map_err(|source| Error::Allocation {
            resource: "added part dependencies",
            source,
        })?;
    for delta in &patch.deltas {
        if delta.after.is_none() {
            removed.insert(folded(delta.name.as_str())?);
        }
        if delta.before.is_none() {
            added.insert(folded(delta.name.as_str())?);
        }
    }
    let root = PackURI::new("/").map_err(Error::Invalid)?;
    for (owner, graph) in std::iter::once((&root, package.rels())).chain(
        package
            .iter_parts()
            .map(|part| (part.partname(), part.rels())),
    ) {
        for relationship in graph.iter().filter(|r| !r.is_external()) {
            let target = folded(relationship.target_partname()?.as_str())?;
            if removed.contains(&target) {
                return Err(Error::PatchConflict);
            }
            if added.contains(&target) && !authorizes_added_target_edge(patch, owner, relationship)
            {
                return Err(Error::PatchConflict);
            }
        }
    }
    Ok(())
}
fn authorizes_added_target_edge(
    patch: &Patch,
    owner: &PackURI,
    edge: &litchi_opc::Relationship,
) -> bool {
    let (before, after) = if owner.as_str() == "/" {
        let Some((a, b)) = &patch.root else {
            return false;
        };
        (Some(a.as_slice()), Some(b.as_slice()))
    } else {
        let index = patch
            .deltas
            .binary_search_by(|d| d.name.as_str().cmp(owner.as_str()))
            .ok()
            .or_else(|| {
                patch
                    .deltas
                    .iter()
                    .position(|d| d.name.is_equivalent_to(owner))
            });
        let Some(index) = index else {
            return false;
        };
        let delta = &patch.deltas[index];
        (
            delta.before.as_ref().map(|s| s.relationships.as_slice()),
            delta.after.as_ref().map(|s| s.relationships.as_slice()),
        )
    };
    fn find<'a>(
        values: Option<&'a [RelationshipState]>,
        id: &str,
    ) -> Option<&'a RelationshipState> {
        let values = values?;
        values
            .binary_search_by(|value| value.id.as_str().cmp(id))
            .ok()
            .map(|i| &values[i])
    }
    let Some(expected) = find(after, edge.r_id()) else {
        return false;
    };
    find(before, edge.r_id()) != Some(expected)
        && expected.kind == edge.reltype()
        && expected.target == edge.target_ref()
        && expected.external == edge.is_external()
}

fn stale() -> Error {
    Error::PatchConflict
}

fn replace_relationships(
    relationships: &mut litchi_opc::Relationships,
    states: &[RelationshipState],
) -> Result<()> {
    let ids: Vec<_> = relationships
        .iter()
        .map(|relationship| relationship.r_id().to_owned())
        .collect();
    for id in ids {
        relationships.remove(&id);
    }
    for state in states {
        relationships.try_add_relationship(
            state.kind.clone(),
            state.target.clone(),
            state.id.clone(),
            if state.external {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )?;
    }
    Ok(())
}

fn require_same_root(left: &Patch, right: &Patch) -> Result<()> {
    if left.presentation_name != right.presentation_name {
        return Err(invalid(
            "opened-presentation patches belong to different roots",
        ));
    }
    Ok(())
}

fn changed_relationships_conflict(
    base: &[RelationshipState],
    left: &[RelationshipState],
    right: &[RelationshipState],
) -> bool {
    let find = |values: &[RelationshipState], id: &str| {
        values
            .binary_search_by(|value| value.id.as_str().cmp(id))
            .ok()
    };
    let conflict = |id: &str| {
        let before = find(base, id).map(|i| &base[i]);
        let a = find(left, id).map(|i| &left[i]);
        let b = find(right, id).map(|i| &right[i]);
        a != before && b != before && a != b
    };
    base.iter()
        .chain(left)
        .chain(right)
        .any(|value| conflict(&value.id))
}
fn deltas_conflict(left: &Delta, right: &Delta) -> bool {
    if left == right {
        return false;
    }
    if left.before != right.before {
        return true;
    }
    match (&left.before, &left.after, &right.after) {
        (Some(base), Some(a), Some(b)) => {
            (a.content_type != base.content_type
                && b.content_type != base.content_type
                && a.content_type != b.content_type)
                || (a.blob != base.blob && b.blob != base.blob && a.blob != b.blob)
                || changed_relationships_conflict(
                    &base.relationships,
                    &a.relationships,
                    &b.relationships,
                )
        },
        _ => true,
    }
}
// Traverse the two sorted outcomes without allocating a merged graph.
fn merged_relationships<'a>(
    base: &'a [RelationshipState],
    left: &'a [RelationshipState],
    right: &'a [RelationshipState],
) -> impl Iterator<Item = &'a RelationshipState> {
    let (mut a, mut b) = (0, 0);
    std::iter::from_fn(move || {
        loop {
            let id = match (left.get(a), right.get(b)) {
                (None, None) => return None,
                (Some(x), None) | (None, Some(x)) => x.id.as_str(),
                (Some(x), Some(y)) => x.id.as_str().min(y.id.as_str()),
            };
            let l = left.get(a).filter(|value| value.id == id);
            let r = right.get(b).filter(|value| value.id == id);
            a += usize::from(l.is_some());
            b += usize::from(r.is_some());
            let previous = base
                .binary_search_by(|value| value.id.as_str().cmp(id))
                .ok()
                .map(|i| &base[i]);
            if let Some(value) = if l == previous { r } else { l } {
                return Some(value);
            }
        }
    })
}
fn merged_relationship_len(
    base: &[RelationshipState],
    left: &[RelationshipState],
    right: &[RelationshipState],
    limits: Limits,
) -> Result<usize> {
    if changed_relationships_conflict(base, left, right) {
        return Err(invalid("patches conflict on a relationship ID"));
    }
    require_relationship_count(base.len(), limits)?;
    let mut count = 0usize;
    let mut length = 4;
    for value in merged_relationships(base, left, right) {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("relationship count overflow"))?;
        require_relationship_count(count, limits)?;
        add_len(
            &mut length,
            relationships_len(std::slice::from_ref(value))? - 4,
        )?;
    }
    Ok(length)
}
fn preflight_join(left: &Patch, right: &Patch, limits: Limits) -> Result<()> {
    if left.conflicts_with(right) {
        return Err(invalid("patches conflict"));
    }
    let count = left
        .deltas
        .len()
        .checked_add(
            right
                .deltas
                .iter()
                .filter(|d| !left.deltas.iter().any(|a| a.name == d.name))
                .count(),
        )
        .ok_or_else(|| invalid("patch part count overflow"))?;
    if count > limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Parts,
            observed: count,
            maximum: limits.max_parts(),
        });
    }
    let mut length = MAGIC.len() + 4 + left.presentation_name.as_str().len() + 1 + 4;
    match (&left.root, &right.root) {
        (Some((base, a)), Some((_, b))) => {
            add_len(&mut length, relationships_len(base)?)?;
            add_len(&mut length, merged_relationship_len(base, a, b, limits)?)?;
        },
        (Some((a, b)), None) | (None, Some((a, b))) => {
            require_relationship_count(a.len(), limits)?;
            require_relationship_count(b.len(), limits)?;
            add_len(&mut length, relationships_len(a)?)?;
            add_len(&mut length, relationships_len(b)?)?;
        },
        (None, None) => {},
    }
    for delta in left.deltas.iter().chain(
        right
            .deltas
            .iter()
            .filter(|d| !left.deltas.iter().any(|a| a.name == d.name)),
    ) {
        add_len(&mut length, 4 + delta.name.as_str().len())?;
        for state in [&delta.before, &delta.after].into_iter().flatten() {
            require_relationship_count(state.relationships.len(), limits)?;
        }
        add_len(&mut length, optional_state_len(delta.before.as_ref())?)?;
        let other = right
            .deltas
            .iter()
            .find(|d| d.name == delta.name && *d != delta);
        if let (Some(other), Some(base), Some(a)) = (other, &delta.before, &delta.after) {
            let b = other
                .after
                .as_ref()
                .ok_or_else(|| invalid("incompatible resource outcomes"))?;
            let content_type = if a.content_type == base.content_type {
                &b.content_type
            } else {
                &a.content_type
            };
            let blob = if a.blob == base.blob {
                &b.blob
            } else {
                &a.blob
            };
            add_len(&mut length, 1 + 4 + content_type.len() + 8)?;
            add_len(&mut length, blob.len())?;
            add_len(
                &mut length,
                merged_relationship_len(
                    &base.relationships,
                    &a.relationships,
                    &b.relationships,
                    limits,
                )?,
            )?;
        } else {
            add_len(&mut length, optional_state_len(delta.after.as_ref())?)?;
        }
        require_capture_bytes(length, limits)?;
    }
    add_len(
        &mut length,
        left.metadata.joined_len(&right.metadata, limits)?,
    )?;
    require_capture_bytes(length, limits)
}
fn merge_relationship_states(
    base: &[RelationshipState],
    left: &[RelationshipState],
    right: &[RelationshipState],
) -> Result<Vec<RelationshipState>> {
    if changed_relationships_conflict(base, left, right) {
        return Err(invalid("patches conflict on a relationship ID"));
    }
    let count = merged_relationships(base, left, right).count();
    let mut merged = Vec::new();
    merged
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation merged relationships",
            source,
        })?;
    merged.extend(merged_relationships(base, left, right).cloned());
    Ok(merged)
}
fn merge_delta(left: &Delta, right: &Delta) -> Result<Delta> {
    if left == right {
        return Ok(left.clone());
    }
    if deltas_conflict(left, right) {
        return Err(invalid(
            "opened-presentation patches conflict on a physical part",
        ));
    }
    let (Some(base), Some(a), Some(b)) = (&left.before, &left.after, &right.after) else {
        return Err(invalid("incompatible resource outcomes"));
    };
    Ok(Delta {
        name: left.name.clone(),
        before: left.before.clone(),
        after: Some(ResourceState {
            content_type: if a.content_type == base.content_type {
                b.content_type.clone()
            } else {
                a.content_type.clone()
            },
            blob: if a.blob == base.blob {
                Arc::clone(&b.blob)
            } else {
                Arc::clone(&a.blob)
            },
            relationships: merge_relationship_states(
                &base.relationships,
                &a.relationships,
                &b.relationships,
            )?,
        }),
    })
}

fn root_conflicts(
    left: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
    right: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
) -> bool {
    matches!((left, right), (Some(left), Some(right)) if left.0 != right.0 || changed_relationships_conflict(&left.0, &left.1, &right.1))
}

fn merge_root(
    left: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
    right: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
) -> Result<Option<(Vec<RelationshipState>, Vec<RelationshipState>)>> {
    match (left, right) {
        (None, None) => Ok(None),
        (Some(value), None) | (None, Some(value)) => Ok(Some(value.clone())),
        (Some(left), Some(right)) if left == right => Ok(Some(left.clone())),
        (Some(left), Some(right)) if left.0 == right.0 => Ok(Some((
            left.0.clone(),
            merge_relationship_states(&left.0, &left.1, &right.1)?,
        ))),
        (Some(_), Some(_)) => Err(invalid(
            "opened-presentation patches have different package-root before-states",
        )),
    }
}

#[cfg(test)]
#[test]
fn mce_retention_intersection_preserves_the_tighter_policy_in_both_orders() -> Result<()> {
    for maximum in [0, 1, 1024, usize::MAX] {
        let left = Limits::default().with_max_retained_mce_bytes(maximum);
        for other in [0, 512, 1024 * 1024, usize::MAX] {
            let right = Limits::default().with_max_retained_mce_bytes(other);
            let expected = Limits::default().with_max_retained_mce_bytes(maximum.min(other));
            assert_eq!(intersect_limits(left, right)?, expected);
            assert_eq!(intersect_limits(right, left)?, expected);
        }
    }
    Ok(())
}

fn intersect_limits(left: Limits, right: Limits) -> Result<Limits> {
    Limits::new(
        left.max_parts().min(right.max_parts()),
        left.max_patch_bytes().min(right.max_patch_bytes()),
        left.max_text_bytes().min(right.max_text_bytes()),
        left.max_history_entries().min(right.max_history_entries()),
        left.max_history_bytes().min(right.max_history_bytes()),
        left.max_retained_candidate_bytes()
            .min(right.max_retained_candidate_bytes()),
    )
    .map(|limits| {
        limits.with_max_retained_mce_bytes(
            left.max_retained_mce_bytes()
                .min(right.max_retained_mce_bytes()),
        )
    })
    .ok_or_else(|| invalid("opened-presentation patch limits are invalid"))
}

fn put_optional_relationship_delta(
    output: &mut Vec<u8>,
    value: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
) -> Result<()> {
    output.push(u8::from(value.is_some()));
    if let Some((before, after)) = value {
        put_relationships(output, before)?;
        put_relationships(output, after)?;
    }
    Ok(())
}

fn put_optional_state(output: &mut Vec<u8>, value: Option<&ResourceState>) -> Result<()> {
    output.push(u8::from(value.is_some()));
    if let Some(value) = value {
        put_bytes32(output, value.content_type.as_bytes())?;
        put_bytes64(output, &value.blob)?;
        put_relationships(output, &value.relationships)?;
    }
    Ok(())
}

fn put_relationships(output: &mut Vec<u8>, relationships: &[RelationshipState]) -> Result<()> {
    put_u32(output, relationships.len())?;
    for relationship in relationships {
        put_bytes32(output, relationship.id.as_bytes())?;
        put_bytes32(output, relationship.kind.as_bytes())?;
        put_bytes32(output, relationship.target.as_bytes())?;
        output.push(u8::from(relationship.external));
    }
    Ok(())
}

// Complete framing/count admission precedes every decoded payload or metadata
// allocation. XML/native-key validation follows under the same finite policy.
fn preflight_wire(bytes: &[u8], limits: Limits) -> Result<()> {
    fn relationships(input: &mut Input<'_>, limits: Limits) -> Result<()> {
        let count = input.usize32()?;
        require_relationship_count(count, limits)?;
        for _ in 0..count {
            input.bytes32()?;
            input.bytes32()?;
            input.bytes32()?;
            if input.byte()? > 1 {
                return Err(invalid("invalid relationship mode"));
            }
        }
        Ok(())
    }
    fn state(input: &mut Input<'_>, limits: Limits) -> Result<()> {
        match input.byte()? {
            0 => Ok(()),
            1 => {
                input.bytes32()?;
                input.bytes64(limits.max_patch_bytes())?;
                relationships(input, limits)
            },
            _ => Err(invalid("invalid resource state marker")),
        }
    }
    let mut input = Input::new(bytes);
    if input.take(MAGIC.len())? != MAGIC {
        return Err(invalid(
            "opened-presentation durable patch has an unsupported version",
        ));
    }
    input.bytes32()?;
    match input.byte()? {
        0 => {},
        1 => {
            relationships(&mut input, limits)?;
            relationships(&mut input, limits)?;
        },
        _ => return Err(invalid("invalid root relationship marker")),
    }
    let count = input.usize32()?;
    if count > limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Parts,
            observed: count,
            maximum: limits.max_parts(),
        });
    }
    for _ in 0..count {
        input.bytes32()?;
        state(&mut input, limits)?;
        state(&mut input, limits)?;
    }
    Metadata::preflight(&mut input, limits)?;
    if !input.is_empty() {
        return Err(invalid(
            "opened-presentation durable patch has trailing bytes",
        ));
    }
    Ok(())
}

fn parse_optional_relationship_delta(
    input: &mut Input<'_>,
    limits: Limits,
) -> Result<Option<(Vec<RelationshipState>, Vec<RelationshipState>)>> {
    match input.byte()? {
        0 => Ok(None),
        1 => Ok(Some((
            parse_relationships(input, limits)?,
            parse_relationships(input, limits)?,
        ))),
        _ => Err(invalid(
            "opened-presentation optional relationship marker is invalid",
        )),
    }
}

fn parse_optional_state(input: &mut Input<'_>, limits: Limits) -> Result<Option<ResourceState>> {
    match input.byte()? {
        0 => Ok(None),
        1 => Ok(Some(ResourceState {
            content_type: parse_text(input.bytes32()?, "content type")?,
            blob: Arc::new(input.bytes64(limits.max_patch_bytes())?.to_vec()),
            relationships: parse_relationships(input, limits)?,
        })),
        _ => Err(invalid(
            "opened-presentation optional resource marker is invalid",
        )),
    }
}

fn parse_relationships(input: &mut Input<'_>, limits: Limits) -> Result<Vec<RelationshipState>> {
    let count = input.usize32()?;
    if count > limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Parts,
            observed: count,
            maximum: limits.max_parts(),
        });
    }
    let mut relationships = Vec::new();
    relationships
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation decoded relationships",
            source,
        })?;
    for _ in 0..count {
        relationships.push(RelationshipState {
            id: parse_text(input.bytes32()?, "relationship ID")?,
            kind: parse_text(input.bytes32()?, "relationship type")?,
            target: parse_text(input.bytes32()?, "relationship target")?,
            external: match input.byte()? {
                0 => false,
                1 => true,
                _ => {
                    return Err(invalid(
                        "opened-presentation relationship target mode is invalid",
                    ));
                },
            },
        });
    }
    relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    if relationships
        .windows(2)
        .any(|pair| pair[0].id == pair[1].id)
    {
        return Err(invalid(
            "opened-presentation patch contains duplicate relationship IDs",
        ));
    }
    Ok(relationships)
}

fn put_u32(output: &mut Vec<u8>, value: usize) -> Result<()> {
    output.extend_from_slice(
        &u32::try_from(value)
            .map_err(|_err| invalid("opened-presentation durable length exceeds u32"))?
            .to_le_bytes(),
    );
    Ok(())
}

fn put_bytes32(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    put_u32(output, value.len())?;
    output.extend_from_slice(value);
    Ok(())
}

fn put_bytes64(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    output.extend_from_slice(
        &u64::try_from(value.len())
            .map_err(|_err| invalid("opened-presentation durable length exceeds u64"))?
            .to_le_bytes(),
    );
    output.extend_from_slice(value);
    Ok(())
}

fn parse_part_name(value: &[u8]) -> Result<PackURI> {
    PackURI::new(parse_text(value, "part name")?)
        .map_err(|_error| invalid("opened-presentation patch part name is invalid"))
}

fn parse_text(value: &[u8], label: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|_err| invalid(format!("opened-presentation patch {label} is not UTF-8")))
}

struct Input<'a> {
    bytes: &'a [u8],
    position: usize,
}

impl<'a> Input<'a> {
    const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes, position: 0 }
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(length)
            .ok_or_else(|| invalid("opened-presentation patch offset overflow"))?;
        let value = self
            .bytes
            .get(self.position..end)
            .ok_or_else(|| invalid("opened-presentation durable patch is truncated"))?;
        self.position = end;
        Ok(value)
    }

    fn byte(&mut self) -> Result<u8> {
        self.take(1)?
            .first()
            .copied()
            .ok_or_else(|| invalid("opened-presentation durable patch is truncated"))
    }

    fn usize32(&mut self) -> Result<usize> {
        let raw: [u8; 4] = self
            .take(4)?
            .try_into()
            .map_err(|_err| invalid("opened-presentation u32 is malformed"))?;
        usize::try_from(u32::from_le_bytes(raw))
            .map_err(|_err| invalid("opened-presentation u32 exceeds usize"))
    }

    fn usize64(&mut self) -> Result<usize> {
        let raw: [u8; 8] = self
            .take(8)?
            .try_into()
            .map_err(|_err| invalid("opened-presentation u64 is malformed"))?;
        usize::try_from(u64::from_le_bytes(raw))
            .map_err(|_err| invalid("opened-presentation u64 exceeds usize"))
    }

    fn bytes32(&mut self) -> Result<&'a [u8]> {
        let length = self.usize32()?;
        self.take(length)
    }

    fn bytes64(&mut self, limit: usize) -> Result<&'a [u8]> {
        let length = self.usize64()?;
        if length > limit {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PayloadBytes),
                observed: length,
                maximum: limit,
            });
        }
        self.take(length)
    }

    const fn is_empty(&self) -> bool {
        self.position == self.bytes.len()
    }
}

#[cfg(test)]
#[path = "patch_capture_tests.rs"]
mod capture_tests;

impl std::fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("resource_count", &self.resource_count())
            .finish_non_exhaustive()
    }
}
impl std::fmt::Debug for Delta {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Delta")
            .field("before_present", &self.before.is_some())
            .field("after_present", &self.after.is_some())
            .finish_non_exhaustive()
    }
}
impl std::fmt::Debug for Conflict {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Conflict").finish_non_exhaustive()
    }
}

#[cfg(test)]
#[path = "patch_metadata_tests.rs"]
mod metadata_tests;
