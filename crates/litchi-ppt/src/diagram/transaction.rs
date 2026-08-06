//! Source-preserving transactions for native diagram BuildList metadata.
//!
//! The editor owns only the fixed metadata fields that this contextual facade
//! can validate: `DiagramBuildEnum` and `BuildAtom.shapeIdRef`. BuildList
//! topology, build identifiers, flags, OfficeArt payloads, timing references,
//! and diagram layout remain inert. Every successful commit carries a patch
//! whose source bytes and source drawing graph are checked before application.

use std::sync::Arc;

use crate::animation::diagram_build::BuildType;
use crate::package::{Error, Result};
use crate::records::Record;

use super::codec;
use super::model::{Build, EditLimits, Id, Inventory, LocatedBuild, ShapeRef};
use super::validation;

/// A deterministic identity of one exact BuildList/drawing source pair.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    /// Returns the compact source fingerprint.
    pub const fn value(self) -> u64 {
        self.0
    }

    fn from_sources(build_list: &[u8], drawing: &[u8]) -> Self {
        let mut hash = 0xcbf2_9ce4_8422_2325u64;
        hash = mix(hash, build_list.len() as u64);
        for byte in build_list {
            hash = mix(hash, u64::from(*byte));
        }
        hash = mix(hash, drawing.len() as u64);
        for byte in drawing {
            hash = mix(hash, u64::from(*byte));
        }
        Self(hash)
    }
}

fn mix(mut hash: u64, value: u64) -> u64 {
    hash ^= value;
    hash.wrapping_mul(0x1000_0000_01b3)
}

/// One semantic operation represented by a diagram metadata patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Change {
    /// Change the fixed-width diagram build mode.
    Mode {
        /// Build identity selected when this operation was staged.
        id: Id,
        /// Source mode.
        before: BuildType,
        /// Committed mode.
        after: BuildType,
    },
    /// Change the referenced OfficeArt shape while retaining the build ID.
    Shape {
        /// Build ID whose shape reference changed.
        build_id: u32,
        /// Complete source identity.
        before: Id,
        /// Complete target identity.
        after: Id,
    },
}

impl Change {
    fn inverse(self) -> Self {
        match self {
            Self::Mode { id, before, after } => Self::Mode {
                id,
                before: after,
                after: before,
            },
            Self::Shape {
                build_id,
                before,
                after,
            } => Self::Shape {
                build_id,
                before: after,
                after: before,
            },
        }
    }
}

/// Immutable, bounded source state for one BuildList and its OfficeArt graph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    build_list: Arc<[u8]>,
    drawing: Arc<[u8]>,
    shape_ids: Arc<[u32]>,
    builds: Vec<LocatedBuild>,
    revision: Revision,
    limits: EditLimits,
}

impl Snapshot {
    /// Parse a complete BuildList and the drawing used to validate its shape
    /// references, retaining both source byte sequences exactly.
    pub fn parse(build_list: impl AsRef<[u8]>, drawing: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(build_list, drawing, EditLimits::default())
    }

    /// Parse with explicit source and graph resource ceilings.
    pub fn parse_with_limits(
        build_list: impl AsRef<[u8]>,
        drawing: impl AsRef<[u8]>,
        limits: EditLimits,
    ) -> Result<Self> {
        validate_limits(limits)?;
        let build_list = build_list.as_ref();
        let drawing = drawing.as_ref();
        if build_list.len() > limits.max_build_list_bytes {
            return invalid("BuildList exceeds the configured transaction byte limit");
        }
        if drawing.len() > limits.max_drawing_bytes {
            return invalid("OfficeArt drawing exceeds the configured transaction byte limit");
        }

        let parsed_drawing = crate::odraw::parse_drawing(drawing)?;
        let mut shape_ids = validation::shape_ids(&parsed_drawing, limits.max_shapes)?;
        shape_ids.sort_unstable();
        let build_list = Arc::<[u8]>::from(build_list.to_vec().into_boxed_slice());
        let drawing = Arc::<[u8]>::from(drawing.to_vec().into_boxed_slice());
        let shape_ids = Arc::<[u32]>::from(shape_ids.into_boxed_slice());
        Self::from_graph(build_list, drawing, shape_ids, limits)
    }

    fn from_graph(
        build_list: Arc<[u8]>,
        drawing: Arc<[u8]>,
        shape_ids: Arc<[u32]>,
        limits: EditLimits,
    ) -> Result<Self> {
        let builds = codec::parse_entries(&build_list, &shape_ids, limits)?;
        let revision = Revision::from_sources(&build_list, &drawing);
        Ok(Self {
            build_list,
            drawing,
            shape_ids,
            builds,
            revision,
            limits,
        })
    }

    /// Exact serialized BuildList bytes.
    pub fn bytes(&self) -> &[u8] {
        &self.build_list
    }

    /// Exact OfficeArt drawing bytes used by shape-reference validation.
    pub fn drawing(&self) -> &[u8] {
        &self.drawing
    }

    /// Source fingerprint, checked in addition to exact bytes by patches.
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Resource ceilings retained by this source owner.
    pub const fn limits(&self) -> EditLimits {
        self.limits
    }

    /// Number of typed diagram builds in the source BuildList.
    pub fn len(&self) -> usize {
        self.builds.len()
    }

    /// Whether the source BuildList contains no typed diagram builds.
    pub fn is_empty(&self) -> bool {
        self.builds.is_empty()
    }

    /// Iterate typed builds without allocating a second public model.
    pub fn builds(&self) -> impl ExactSizeIterator<Item = Build> + '_ {
        self.builds.iter().map(|entry| entry.build)
    }

    /// Find one typed build by its checked `(buildId, shapeIdRef)` identity.
    pub fn get(&self, id: Id) -> Option<Build> {
        self.builds
            .iter()
            .find(|entry| entry.build.id() == id)
            .map(|entry| entry.build)
    }

    /// Recreate the existing read-only inventory projection on demand.
    pub fn inventory(&self) -> Result<Inventory<'_>> {
        let (record, consumed) = Record::parse_strict(&self.build_list, 0)?;
        if consumed != self.build_list.len() {
            return Err(Error::Corrupted("BuildList has trailing bytes".to_string()));
        }
        Inventory::parse_with_limits(
            &record,
            &self.drawing,
            super::model::Limits {
                max_diagrams: self.limits.max_diagrams,
                max_shapes_per_diagram: self.limits.max_shapes,
                max_payloads_per_diagram: self.limits.max_build_list_bytes,
            },
        )
    }

    /// Start an isolated edit over this exact source.
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            candidate: self.build_list.to_vec(),
            builds: self.builds.clone(),
            changes: Vec::new(),
        }
    }
}

/// Failure-atomic semantic editor over one diagram BuildList snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Vec<u8>,
    builds: Vec<LocatedBuild>,
    changes: Vec<Change>,
}

impl Transaction {
    /// Borrow the immutable source snapshot.
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Iterate the currently staged typed builds.
    pub fn builds(&self) -> impl ExactSizeIterator<Item = Build> + '_ {
        self.builds.iter().map(|entry| entry.build)
    }

    /// Whether staged bytes differ from the source bytes.
    pub fn is_changed(&self) -> bool {
        self.candidate.as_slice() != self.source.bytes()
    }

    /// Borrow ordered semantic changes staged so far.
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Change one diagram's checked build mode while retaining all other
    /// fields and records, including unknown enum values not selected here.
    pub fn set_mode(&mut self, id: Id, mode: BuildType) -> Result<()> {
        let entry = self.entry(id)?;
        let before = entry.build.mode();
        if before == mode {
            return Ok(());
        }
        let mut candidate = self.candidate.clone();
        codec::rewrite_mode(&mut candidate, entry.offset, mode)?;
        let builds = codec::parse_entries(&candidate, &self.source.shape_ids, self.source.limits)?;
        self.candidate = candidate;
        self.builds = builds;
        self.changes.push(Change::Mode {
            id,
            before,
            after: mode,
        });
        Ok(())
    }

    /// Change only a diagram's `shapeIdRef` to an existing OfficeArt shape.
    ///
    /// Build IDs, timing references, list order, and OfficeArt payloads are
    /// intentionally not editable by this owner.
    pub fn set_shape_id(&mut self, id: Id, shape_id: u32) -> Result<Id> {
        if self.source.shape_ids.binary_search(&shape_id).is_err() {
            return invalid("diagram shape identity does not exist in the OfficeArt graph");
        }
        let entry = self.entry(id)?;
        if id.shape_id() == shape_id {
            return Ok(id);
        }
        let target = Id::new(id.build_id(), shape_id);
        if self.builds.iter().any(|other| other.build.id() == target) {
            return invalid("diagram shape edit would create a duplicate build identity");
        }
        let mut candidate = self.candidate.clone();
        codec::rewrite_shape_id(&mut candidate, entry.offset, shape_id)?;
        let builds = codec::parse_entries(&candidate, &self.source.shape_ids, self.source.limits)?;
        self.candidate = candidate;
        self.builds = builds;
        self.changes.push(Change::Shape {
            build_id: id.build_id(),
            before: id,
            after: target,
        });
        Ok(target)
    }

    /// Checked shape-reference spelling for callers already holding a
    /// contextual [`ShapeRef`].
    pub fn set_shape(&mut self, id: Id, shape: ShapeRef) -> Result<Id> {
        self.set_shape_id(id, shape.id())
    }

    /// Capture the current candidate without publishing a patch.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if self.candidate.as_slice() == self.source.bytes() {
            return Ok(self.source.clone());
        }
        let bytes = Arc::<[u8]>::from(self.candidate.clone().into_boxed_slice());
        Snapshot::from_graph(
            bytes,
            self.source.drawing.clone(),
            self.source.shape_ids.clone(),
            self.source.limits,
        )
    }

    /// Validate and publish the candidate with its reversible source patch.
    pub fn commit(self) -> Result<Commit> {
        let unchanged = self.candidate.as_slice() == self.source.bytes();
        let snapshot = if unchanged {
            self.source.clone()
        } else {
            let bytes = Arc::<[u8]>::from(self.candidate.into_boxed_slice());
            Snapshot::from_graph(
                bytes,
                self.source.drawing.clone(),
                self.source.shape_ids.clone(),
                self.source.limits,
            )?
        };
        let changes = if unchanged { Vec::new() } else { self.changes };
        let patch = Patch {
            base: self.source.revision,
            target: snapshot.revision,
            before: self.source.build_list.clone(),
            after: snapshot.build_list.clone(),
            drawing: self.source.drawing.clone(),
            changes,
            limits: self.source.limits,
        };
        Ok(Commit { snapshot, patch })
    }

    /// Alias for move-owned writer terminology.
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }

    /// Discard staged edits and recover the source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn entry(&self, id: Id) -> Result<LocatedBuild> {
        self.builds
            .iter()
            .find(|entry| entry.build.id() == id)
            .copied()
            .ok_or_else(|| Error::InvalidFormat("diagram build identity was not found".into()))
    }
}

/// A successful target snapshot and its source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Published target snapshot.
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible patch from the source to the target.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split the target and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Reversible, source-checked BuildList byte patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    drawing: Arc<[u8]>,
    changes: Vec<Change>,
    limits: EditLimits,
}

impl Patch {
    /// Source revision required by [`Self::apply`].
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Target revision produced by [`Self::apply`].
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Ordered semantic operations represented by this patch.
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Exact source BuildList bytes.
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact target BuildList bytes.
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether this patch is a semantic and byte-level no-op.
    pub const fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Apply only to the exact source snapshot used to create this patch.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision != self.base
            || current.build_list.as_ref() != self.before.as_ref()
            || current.drawing.as_ref() != self.drawing.as_ref()
        {
            return invalid("diagram patch source does not match its base snapshot");
        }
        if self.before.as_ref() == self.after.as_ref() {
            return Ok(current.clone());
        }
        Snapshot::from_graph(
            self.after.clone(),
            current.drawing.clone(),
            current.shape_ids.clone(),
            self.limits,
        )
    }

    /// Apply the inverse to the exact committed target snapshot.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.inverse().apply(current)
    }

    /// Reapply this patch to its exact source snapshot.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.apply(current)
    }

    /// Build a source-checked inverse patch.
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            drawing: self.drawing.clone(),
            changes: self
                .changes
                .iter()
                .rev()
                .copied()
                .map(Change::inverse)
                .collect(),
            limits: self.limits,
        }
    }
}

fn validate_limits(limits: EditLimits) -> Result<()> {
    if limits.max_build_list_bytes < 8 {
        return invalid("diagram transaction BuildList limit must include a record header");
    }
    Ok(())
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
