//! Package-level slide-library synchronization edits (record 0357).
//!
//! [`Package::edit_slide_sync`] stages the optional
//! `RoundTripSlideSyncInfo12` container of one slide, [`Edit::commit`]
//! publishes a new [`Package`] through one append-only user edit that
//! rewrites only that slide, and [`Package::apply_slide_sync`] replays a
//! [`Patch`] onto the exact package it was made from. Dropping an [`Edit`]
//! discards it. The library URL stays inert: nothing is contacted.

use litchi_core::Position;

use super::model::Limits;
use super::transaction::RecordEdit;
use super::{Change, Commit, Patch, Result, Snapshot, Synchronization};
use crate::Package;
use crate::scoped::{self, Base};
use crate::slide_order::{self, records::Owner};

/// A staged edit of one slide's synchronization metadata over one immutable
/// [`Package`].
pub struct Edit {
    base: Base,
    native: slide_order::Snapshot,
    slide: Position,
    before: Vec<u8>,
    inner: RecordEdit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.slide)
            .field("changes", &self.inner.changes().len())
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(package: &Package, slide: Position) -> Result<Self> {
        let base = Base::capture(package)?;
        let native = base.native()?;
        let before = base.owner_record(&native, Owner::Slide(slide))?;
        let inner = Snapshot::parse_with_limits(&before, Limits::default())
            .map_err(scoped::publication("edit_slide_sync"))?
            .edit();
        Ok(Self {
            base,
            native,
            slide,
            before,
            inner,
        })
    }

    /// The slide whose synchronization metadata is staged.
    #[must_use]
    pub const fn slide(&self) -> Position {
        self.slide
    }

    /// The staged synchronization metadata, or `None` when the slide has
    /// none.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the staged record cannot be
    /// read as synchronization metadata.
    pub fn synchronization(&self) -> Result<Option<Synchronization>> {
        Ok(self
            .inner
            .synchronization()
            .map_err(scoped::publication("slide_sync.synchronization"))?)
    }

    /// The staged operations, in order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        self.inner.changes()
    }

    /// Set or replace the slide's synchronization metadata. An equal value
    /// is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the record grammar refuses the
    /// value, leaving the edit unchanged.
    pub fn set(&mut self, value: Synchronization) -> Result<()> {
        Ok(self
            .inner
            .set(value)
            .map_err(scoped::staging("slide_sync.set"))?)
    }

    /// Remove the slide's synchronization metadata, returning whether it was
    /// present.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the slide's records cannot be
    /// edited, leaving the edit unchanged.
    pub fn clear(&mut self) -> Result<bool> {
        Ok(self
            .inner
            .clear()
            .map_err(scoped::staging("slide_sync.clear"))?)
    }

    /// Publish the staged metadata as a new package. An edit with no net
    /// change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Signed`] for a changed write to a signed
    /// package and [`super::Error::Package`] when the result cannot be
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            base,
            native,
            slide,
            before,
            inner,
        } = self;
        if inner.changes().is_empty() {
            return Ok(super::published(base.unchanged()));
        }
        let (target, _) = inner
            .commit()
            .map_err(scoped::publication("slide_sync.commit"))?
            .into_parts();
        Ok(super::published(base.publish_owner(
            &native,
            Owner::Slide(slide),
            before,
            target.bytes().to_vec(),
        )?))
    }
}

impl Package {
    /// Start an edit of the synchronization metadata of the slide at a
    /// zero-based position.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no slide is at
    /// `slide`, and [`super::Error::Package`] for an encrypted or unreadable
    /// presentation.
    pub fn edit_slide_sync(&self, slide: Position) -> Result<Edit> {
        Edit::new(self, slide)
    }

    /// Apply a slide-sync patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_slide_sync(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
