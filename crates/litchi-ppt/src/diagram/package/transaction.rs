//! Failure-atomic publication transactions for one owning slide payload.

use std::sync::Arc;

use crate::animation::diagram_build::BuildType;
use crate::package::{Error, Result};

use super::super::model::{Build, Id, ShapeRef};
use super::super::transaction::{Change, Patch as DiagramPatch, Transaction as DiagramTransaction};
use super::codec;
use super::model::{SlideLimits, SlideRevision, SlideSnapshot};

/// An isolated editor over the supported diagram metadata inside one slide.
#[derive(Debug, Clone)]
pub struct SlideEditor {
    source: SlideSnapshot,
    diagram: DiagramTransaction,
}

impl SlideEditor {
    pub(super) fn open(source: SlideSnapshot) -> Self {
        Self {
            diagram: source.diagram.edit(),
            source,
        }
    }

    /// Borrow the exact source slide snapshot.
    #[must_use]
    pub const fn source(&self) -> &SlideSnapshot {
        &self.source
    }

    /// Iterate the currently staged typed diagram builds.
    #[must_use]
    pub fn builds(&self) -> impl ExactSizeIterator<Item = Build> + '_ {
        self.diagram.builds()
    }

    /// Whether supported metadata differs from the source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.diagram.is_changed()
    }

    /// Borrow semantic diagram changes staged in this editor.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        self.diagram.changes()
    }

    /// Change one diagram's validated `DiagramBuildEnum` value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_mode(&mut self, id: Id, mode: BuildType) -> Result<()> {
        self.diagram.set_mode(id, mode)
    }

    /// Change one diagram's `shapeIdRef` to an existing `OfficeArt` shape.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_shape_id(&mut self, id: Id, shape_id: u32) -> Result<Id> {
        self.diagram.set_shape_id(id, shape_id)
    }

    /// Checked shape-reference spelling for callers holding a contextual
    /// [`ShapeRef`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_shape(&mut self, id: Id, shape: ShapeRef) -> Result<Id> {
        self.diagram.set_shape(id, shape)
    }

    /// Capture the candidate slide without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn snapshot(&self) -> Result<SlideSnapshot> {
        let diagram = self.diagram.snapshot()?;
        let bytes = codec::replace_build_list(
            self.source.bytes(),
            self.source.build_range.clone(),
            diagram.bytes(),
        )?;
        SlideSnapshot::from_bytes_with_limits(bytes, self.source.limits())
    }

    /// Validate and publish a source-checked, reversible slide patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<SlideCommit> {
        let diagram_commit = self.diagram.commit()?;
        let source = self.source;
        let target_diagram = diagram_commit.snapshot();
        if target_diagram.bytes() == source.build_list() {
            let patch = SlidePatch {
                base: source.revision(),
                target: source.revision(),
                before: source.bytes.clone(),
                after: source.bytes.clone(),
                diagram: diagram_commit.patch().clone(),
                limits: source.limits(),
            };
            return Ok(SlideCommit {
                snapshot: source,
                patch,
            });
        }

        let bytes = codec::replace_build_list(
            source.bytes(),
            source.build_range.clone(),
            target_diagram.bytes(),
        )?;
        let snapshot = SlideSnapshot::from_bytes_with_limits(bytes, source.limits())?;
        if snapshot.build_list() != target_diagram.bytes() || snapshot.drawing() != source.drawing()
        {
            return Err(Error::Corrupted(
                "published diagram BuildList failed owning-slide validation".into(),
            ));
        }
        let patch = SlidePatch {
            base: source.revision(),
            target: snapshot.revision(),
            before: source.bytes,
            after: snapshot.bytes.clone(),
            diagram: diagram_commit.patch().clone(),
            limits: snapshot.limits(),
        };
        Ok(SlideCommit { snapshot, patch })
    }

    /// Alias matching move-owned writer terminology.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn finish(self) -> Result<SlideCommit> {
        self.commit()
    }

    /// Discard staged changes and recover the exact source slide.
    #[must_use]
    pub fn rollback(self) -> SlideSnapshot {
        self.source
    }
}

/// A committed owning-slide snapshot and reversible whole-slide patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideCommit {
    snapshot: SlideSnapshot,
    patch: SlidePatch,
}

impl SlideCommit {
    /// Borrow the validated target slide snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &SlideSnapshot {
        &self.snapshot
    }

    /// Borrow the source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &SlidePatch {
        &self.patch
    }

    /// Consume the commit into its target snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (SlideSnapshot, SlidePatch) {
        (self.snapshot, self.patch)
    }

    /// Undo against the exact committed slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &SlideSnapshot) -> Result<SlideSnapshot> {
        self.patch.undo(current)
    }

    /// Redo against the exact source slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &SlideSnapshot) -> Result<SlideSnapshot> {
        self.patch.redo(current)
    }
}

/// A source-checked reversible patch for one complete `SlideContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlidePatch {
    base: SlideRevision,
    target: SlideRevision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    diagram: DiagramPatch,
    limits: SlideLimits,
}

impl SlidePatch {
    /// Source revision required for forward application.
    #[must_use]
    pub const fn base(&self) -> SlideRevision {
        self.base
    }

    /// Target revision produced by forward application.
    #[must_use]
    pub const fn target(&self) -> SlideRevision {
        self.target
    }

    /// Exact source slide bytes bound to this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact replacement slide bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// The nested typed diagram patch represented by this publication.
    #[must_use]
    pub const fn diagram(&self) -> &DiagramPatch {
        &self.diagram
    }

    /// Alias for callers that name the nested owner explicitly.
    #[must_use]
    pub const fn build(&self) -> &DiagramPatch {
        &self.diagram
    }

    /// Whether this patch is an exact no-op.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.diagram.is_empty()
    }

    /// Apply only to the exact source slide used to create this patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn apply(&self, current: &SlideSnapshot) -> Result<SlideSnapshot> {
        if current.revision() != self.base || current.bytes() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "diagram publication patch source does not match its base slide".into(),
            ));
        }
        let diagram = self.diagram.apply(&current.diagram)?;
        if diagram.bytes() != self.diagram.after() {
            return Err(Error::Corrupted(
                "diagram publication patch BuildList target is inconsistent".into(),
            ));
        }
        if self.before.as_ref() == self.after.as_ref() {
            return Ok(current.clone());
        }
        SlideSnapshot::from_bytes_with_limits(self.after.to_vec(), self.limits)
    }

    /// Apply the inverse patch to the exact committed target slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &SlideSnapshot) -> Result<SlideSnapshot> {
        self.inverse().apply(current)
    }

    /// Reapply this patch to its exact source slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &SlideSnapshot) -> Result<SlideSnapshot> {
        self.apply(current)
    }

    /// Build a source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            diagram: self.diagram.inverse(),
            limits: self.limits,
        }
    }
}
