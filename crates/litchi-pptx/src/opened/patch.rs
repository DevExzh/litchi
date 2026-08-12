//! Durable exact-source package patches, deterministic merges, and bounded history.

use std::collections::VecDeque;
use std::sync::Arc;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use super::model::{Limits, Snapshot, capture, invalid};
use crate::{Error, Result};

const MAGIC: &[u8; 8] = b"LPTX0002";

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipState {
    id: String,
    kind: String,
    target: String,
    external: bool,
}

impl RelationshipState {
    fn capture(relationships: &litchi_opc::Relationships) -> Vec<Self> {
        let mut states: Vec<_> = relationships
            .iter()
            .map(|relationship| Self {
                id: relationship.r_id().to_owned(),
                kind: relationship.reltype().to_owned(),
                target: relationship.target_ref().to_owned(),
                external: relationship.is_external(),
            })
            .collect();
        states.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        states
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ResourceState {
    content_type: String,
    blob: Arc<Vec<u8>>,
    relationships: Vec<RelationshipState>,
}

impl ResourceState {
    fn capture(part: &dyn Part) -> Self {
        Self {
            content_type: part.content_type().to_owned(),
            blob: part.blob_arc(),
            relationships: RelationshipState::capture(part.rels()),
        }
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

#[derive(Debug, Clone, PartialEq, Eq)]
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
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Conflict {
    name: PackURI,
}

impl Conflict {
    /// Physical OPC part whose two candidate states differ.
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
    left: Delta,
    right: Delta,
    resolution: Option<Resolution>,
}

/// Immutable, non-publishing plan for a deterministic three-way merge.
#[derive(Debug, Clone)]
pub struct ThreeWayPlan {
    presentation_name: PackURI,
    automatic: Vec<Delta>,
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
        let mut deltas = self.automatic.clone();
        for entry in &self.conflicts {
            match entry.resolution {
                Some(Resolution::Base) => {},
                Some(Resolution::Left) => deltas.push(entry.left.clone()),
                Some(Resolution::Right) => deltas.push(entry.right.clone()),
                None => {
                    return Err(invalid(
                        "opened-presentation three-way merge has unresolved conflicts",
                    ));
                },
            }
        }
        Patch::new(self.presentation_name.clone(), deltas, None, self.limits)
    }
}

/// Durable, exact-source patch over a finite set of OPC resources.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    pub(crate) presentation_name: PackURI,
    pub(crate) deltas: Vec<Delta>,
    root: Option<(Vec<RelationshipState>, Vec<RelationshipState>)>,
    pub(crate) limits: Limits,
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
        let mut names: Vec<_> = source
            .iter_parts()
            .chain(target.iter_parts())
            .map(|part| part.partname().clone())
            .collect();
        names.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
        names.dedup();
        let mut deltas = Vec::new();
        for name in names {
            let before = source.get_part(&name).ok().map(ResourceState::capture);
            let after = target.get_part(&name).ok().map(ResourceState::capture);
            if before != after {
                deltas.push(Delta {
                    name,
                    before,
                    after,
                });
            }
        }
        let before_root = RelationshipState::capture(source.rels());
        let after_root = RelationshipState::capture(target.rels());
        let root = (before_root != after_root).then_some((before_root, after_root));
        Self::new(presentation_name, deltas, root, limits)
    }

    fn new(
        presentation_name: PackURI,
        mut deltas: Vec<Delta>,
        root: Option<(Vec<RelationshipState>, Vec<RelationshipState>)>,
        limits: Limits,
    ) -> Result<Self> {
        deltas.sort_unstable_by(|left, right| left.name.as_str().cmp(right.name.as_str()));
        if deltas.len() > limits.max_parts() {
            return Err(Error::Limit {
                resource: "opened-presentation patch parts",
                limit: limits.max_parts(),
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
        let patch = Self {
            presentation_name,
            deltas,
            root,
            limits,
        };
        patch.require_encoded_limit()?;
        Ok(patch)
    }

    /// Whether this patch changes no OPC resource.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.deltas.is_empty() && self.root.is_none()
    }

    /// Number of exact part and optional package-root resources in the write set.
    #[must_use]
    pub fn resource_count(&self) -> usize {
        self.deltas.len() + usize::from(self.root.is_some())
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

    pub(crate) fn exact_slide_removal_change(&self) -> Option<(SinglePartChange, &PackURI)> {
        if self.root.is_some() || self.deltas.len() != 2 {
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
        let slide = self
            .deltas
            .iter()
            .find(|delta| delta.name != self.presentation_name)?;
        match (&slide.before, &slide.after) {
            (Some(before), None)
                if before.content_type == litchi_opc::constants::content_type::PML_SLIDE =>
            {
                Some((SinglePartChange::Remove, &slide.name))
            },
            (None, Some(after))
                if after.content_type == litchi_opc::constants::content_type::PML_SLIDE =>
            {
                Some((SinglePartChange::Add, &slide.name))
            },
            _ => None,
        }
    }

    pub(crate) fn has_same_changes(&self, other: &Self) -> bool {
        self.presentation_name == other.presentation_name
            && self.deltas == other.deltas
            && self.root == other.root
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
        }
    }

    /// Whether two patches select different outcomes for any physical resource.
    #[must_use]
    pub fn conflicts_with(&self, other: &Self) -> bool {
        root_conflicts(self.root.as_ref(), other.root.as_ref())
            || self.deltas.iter().any(|left| {
                other.deltas.iter().any(|right| {
                    left.name == right.name
                        && (left.before != right.before || left.after != right.after)
                })
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
        let mut deltas = self.deltas.clone();
        for candidate in &other.deltas {
            if let Some(existing) = deltas.iter().find(|delta| delta.name == candidate.name) {
                if existing.before != candidate.before || existing.after != candidate.after {
                    return Err(invalid(
                        "opened-presentation patches conflict on a physical part",
                    ));
                }
            } else {
                deltas.push(candidate.clone());
            }
        }
        let root = merge_root(self.root.as_ref(), other.root.as_ref())?;
        Self::new(self.presentation_name.clone(), deltas, root, limits)
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
        if left.root.is_some() || right.root.is_some() {
            return Err(invalid(
                "opened-presentation three-way planning does not merge package-root relationships",
            ));
        }
        let limits = intersect_limits(left.limits, right.limits)?;
        let mut automatic = Vec::new();
        let mut conflicts = Vec::new();
        for delta in &left.deltas {
            match right.deltas.iter().find(|other| other.name == delta.name) {
                None => automatic.push(delta.clone()),
                Some(other) if delta.after == other.after => automatic.push(delta.clone()),
                Some(other) => conflicts.push(PlannedConflict {
                    conflict: Conflict {
                        name: delta.name.clone(),
                    },
                    left: delta.clone(),
                    right: other.clone(),
                    resolution: None,
                }),
            }
        }
        for delta in &right.deltas {
            if !left.deltas.iter().any(|other| other.name == delta.name) {
                automatic.push(delta.clone());
            }
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
            conflicts,
            limits,
        })
    }

    /// Serialize this patch into the stable `LPTX0002` binary format.
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
        output.extend_from_slice(MAGIC);
        put_bytes32(&mut output, self.presentation_name.as_str().as_bytes())?;
        put_optional_relationship_delta(&mut output, self.root.as_ref())?;
        put_u32(&mut output, self.deltas.len())?;
        for delta in &self.deltas {
            put_bytes32(&mut output, delta.name.as_str().as_bytes())?;
            put_optional_state(&mut output, delta.before.as_ref())?;
            put_optional_state(&mut output, delta.after.as_ref())?;
        }
        if output.len() != length {
            return Err(invalid(
                "opened-presentation durable patch length changed during encoding",
            ));
        }
        Ok(output)
    }

    /// Parse a stable durable patch under conservative finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed or unbounded input.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a stable durable patch under caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed or unbounded input.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        if bytes.len() > limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation durable patch bytes",
                limit: limits.max_patch_bytes(),
            });
        }
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
                resource: "opened-presentation patch parts",
                limit: limits.max_parts(),
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
        if !input.is_empty() {
            return Err(invalid(
                "opened-presentation durable patch has trailing bytes",
            ));
        }
        Self::new(presentation_name, deltas, root, limits)
    }

    pub(crate) fn encoded_len(&self) -> Result<usize> {
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
                    resource: "opened-presentation durable patch bytes",
                    limit: self.limits.max_patch_bytes(),
                });
            }
        }
        Ok(length)
    }

    fn require_encoded_limit(&self) -> Result<()> {
        let length = self.encoded_len()?;
        if length > self.limits.max_patch_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation durable patch bytes",
                limit: self.limits.max_patch_bytes(),
            });
        }
        Ok(())
    }
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
        if patch.is_empty() {
            return Ok(());
        }
        self.clear_redo();
        let length = patch.encoded_len()?;
        if length > self.limits.max_history_bytes() {
            return Err(Error::Limit {
                resource: "opened-presentation history bytes",
                limit: self.limits.max_history_bytes(),
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

pub(crate) fn apply(package: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    apply_with_revision(package, patch, None)
}

pub(crate) fn apply_exact_revision(
    package: &mut OpcPackage,
    patch: &Patch,
    result_revision: [u8; 32],
) -> Result<Snapshot> {
    apply_with_revision(package, patch, Some(result_revision))
}

fn apply_with_revision(
    package: &mut OpcPackage,
    patch: &Patch,
    result_revision: Option<[u8; 32]>,
) -> Result<Snapshot> {
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
    if patch.is_empty() {
        let snapshot = capture(package, patch.limits)?;
        if result_revision.is_some_and(|expected| snapshot.revision() != expected) {
            return Err(invalid(
                "opened-presentation candidate has an unexpected complete-package revision",
            ));
        }
        return Ok(snapshot);
    }
    let mut candidate = package.clone();
    if let Some((_before, after)) = &patch.root {
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
                replace_relationships(part.rels_mut(), &state.relationships)?;
            },
            (None, None) => {
                return Err(invalid(
                    "opened-presentation candidate contains a no-op resource",
                ));
            },
        }
    }
    let snapshot = capture(&candidate, patch.limits)?;
    validate_after(&candidate, patch)?;
    if result_revision.is_some_and(|expected| snapshot.revision() != expected) {
        return Err(invalid(
            "opened-presentation candidate has an unexpected complete-package revision",
        ));
    }
    *package = candidate;
    Ok(snapshot)
}

fn validate_before(package: &OpcPackage, patch: &Patch) -> Result<()> {
    if let Some((before, _after)) = &patch.root
        && RelationshipState::capture(package.rels()) != *before
    {
        return Err(stale());
    }
    for delta in &patch.deltas {
        let current = package
            .get_part(&delta.name)
            .ok()
            .map(ResourceState::capture);
        if current != delta.before {
            return Err(stale());
        }
    }
    Ok(())
}

fn validate_after(package: &OpcPackage, patch: &Patch) -> Result<()> {
    if let Some((_before, after)) = &patch.root
        && RelationshipState::capture(package.rels()) != *after
    {
        return Err(invalid(
            "opened-presentation published root relationships differ from the patch target",
        ));
    }
    for delta in &patch.deltas {
        let current = package
            .get_part(&delta.name)
            .ok()
            .map(ResourceState::capture);
        if current != delta.after {
            return Err(invalid(
                "opened-presentation published resource differs from its patch target",
            ));
        }
    }
    Ok(())
}

fn stale() -> Error {
    Error::UnsafeEdit {
        operation: "apply_opened_presentation_patch",
        reason: "an opened-presentation patch resource is stale",
    }
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

fn root_conflicts(
    left: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
    right: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
) -> bool {
    matches!((left, right), (Some(left), Some(right)) if left != right)
}

fn merge_root(
    left: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
    right: Option<&(Vec<RelationshipState>, Vec<RelationshipState>)>,
) -> Result<Option<(Vec<RelationshipState>, Vec<RelationshipState>)>> {
    match (left, right) {
        (None, None) => Ok(None),
        (Some(value), None) | (None, Some(value)) => Ok(Some(value.clone())),
        (Some(left), Some(right)) if left == right => Ok(Some(left.clone())),
        (Some(_), Some(_)) => Err(invalid(
            "opened-presentation patches conflict on package-root relationships",
        )),
    }
}

fn intersect_limits(left: Limits, right: Limits) -> Result<Limits> {
    Limits::new(
        left.max_parts().min(right.max_parts()),
        left.max_patch_bytes().min(right.max_patch_bytes()),
        left.max_text_bytes().min(right.max_text_bytes()),
        left.max_history_entries().min(right.max_history_entries()),
        left.max_history_bytes().min(right.max_history_bytes()),
    )
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
            resource: "opened-presentation patch relationships",
            limit: limits.max_parts(),
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
    PackURI::new(parse_text(value, "part name")?).map_err(Error::Invalid)
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
                resource: "opened-presentation durable resource bytes",
                limit,
            });
        }
        self.take(length)
    }

    const fn is_empty(&self) -> bool {
        self.position == self.bytes.len()
    }
}
