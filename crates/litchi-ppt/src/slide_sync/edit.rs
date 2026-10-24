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
use crate::scoped::{self, Base, Failure, Lazy};
use crate::slide_order::{self, records::Owner};

/// A staged edit of one slide's synchronization metadata over one immutable
/// [`Package`].
pub struct Edit {
    base: Base,
    slide: Position,
    state: Lazy<State>,
}

struct State {
    native: slide_order::Snapshot,
    before: Vec<u8>,
    inner: RecordEdit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.slide)
            .field("initialized", &self.state.get().is_some())
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(package: &Package, slide: Position) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            base,
            slide,
            state: Lazy::new(),
        })
    }

    fn initialize_state(&self) -> std::result::Result<State, Failure> {
        let native = self.base.native()?;
        let before = self.base.owner_record(&native, Owner::Slide(self.slide))?;
        let inner = Snapshot::parse_with_limits(&before, Limits::default())
            .map_err(scoped::publication("edit_slide_sync"))?
            .edit();
        Ok(State {
            native,
            before,
            inner,
        })
    }

    fn state(&self) -> std::result::Result<&State, Failure> {
        self.state.get_or_try_init(|| self.initialize_state())
    }

    fn state_mut(&mut self) -> std::result::Result<&mut State, Failure> {
        let base = self.base.clone();
        let slide = self.slide;
        self.state.get_or_try_init_mut(|| {
            let native = base.native()?;
            let before = base.owner_record(&native, Owner::Slide(slide))?;
            let inner = Snapshot::parse_with_limits(&before, Limits::default())
                .map_err(scoped::publication("edit_slide_sync"))?
                .edit();
            Ok(State {
                native,
                before,
                inner,
            })
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
            .state()?
            .inner
            .synchronization()
            .map_err(scoped::publication("slide_sync.synchronization"))?)
    }

    /// The staged operations, in order.
    #[must_use = "handle the fallible slide-sync changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.state()?.inner.changes())
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
            .state_mut()?
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
            .state_mut()?
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
        self.state
            .get_or_try_init(|| self.initialize_state())
            .map_err(super::Error::from)?;
        let Self { base, slide, state } = self;
        let State {
            native,
            before,
            inner,
        } = state.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT slide-sync edit state was not initialized".into(),
            ))
        })?;
        if inner.changes().is_empty() {
            return Ok(super::published(base.unchanged()?));
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
    /// An encrypted presentation is refused immediately. A missing slide or
    /// unreadable record is reported by the first projection, staging
    /// operation, or commit.
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
