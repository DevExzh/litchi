//! Package-level diagram metadata edits (record 0357).
//!
//! [`Package::edit_diagrams`] stages edits of the fixed-width diagram build
//! metadata of one slide (the `BuildList` in its `___PPT10` programmable tag,
//! validated against the slide's `PPDrawing`), [`Edit::commit`] publishes a
//! new [`Package`] through one append-only user edit that rewrites only that
//! slide, and [`Package::apply_diagrams`] replays a [`Patch`] onto the exact
//! package it was made from. Dropping an [`Edit`] discards it.

use litchi_core::Position;

use super::package::SlideEditor;
use super::{Build, Change, Commit, Id, Patch, Result, ShapeRef, SlideLimits, SlideSnapshot};
use crate::Package;
use crate::animation::diagram_build::BuildType;
use crate::package::Error as PackageError;
use crate::scoped::{self, Base, Failure, Lazy};
use crate::slide_order::{self, records::Owner};

const OPERATION: &str = "edit_diagrams";

/// A staged edit of one slide's diagram build metadata over one immutable
/// [`Package`].
pub struct Edit {
    base: Base,
    slide: Position,
    state: Lazy<State>,
}

struct State {
    native: slide_order::Snapshot,
    before: Vec<u8>,
    inner: SlideEditor,
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
        let snapshot =
            SlideSnapshot::from_bytes_with_limits(before.clone(), SlideLimits::default()).map_err(
                |error| match error {
                    PackageError::InvalidFormat(_) => scoped::unsupported_target(
                        OPERATION,
                        "the slide carries no editable diagram build list",
                    ),
                    error => Failure::Package(error),
                },
            )?;
        Ok(State {
            native,
            before,
            inner: snapshot.edit(),
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
            let snapshot =
                SlideSnapshot::from_bytes_with_limits(before.clone(), SlideLimits::default())
                    .map_err(|error| match error {
                        PackageError::InvalidFormat(_) => scoped::unsupported_target(
                            OPERATION,
                            "the slide carries no editable diagram build list",
                        ),
                        error => Failure::Package(error),
                    })?;
            Ok(State {
                native,
                before,
                inner: snapshot.edit(),
            })
        })
    }

    fn require_build(&self, operation: &'static str, id: Id) -> Result<()> {
        if self.state()?.inner.builds().any(|build| build.id() == id) {
            Ok(())
        } else {
            Err(scoped::not_found(
                operation,
                "the slide has no diagram build with this identity",
            )
            .into())
        }
    }

    /// The slide whose diagram metadata is staged.
    #[must_use]
    pub const fn slide(&self) -> Position {
        self.slide
    }

    /// The staged typed diagram builds, in `BuildList` order.
    #[must_use = "handle the fallible diagram projection result"]
    pub fn builds(&self) -> Result<impl ExactSizeIterator<Item = Build> + '_> {
        Ok(self.state()?.inner.builds())
    }

    /// The staged operations, in order.
    #[must_use = "handle the fallible diagram changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.state()?.inner.changes())
    }

    /// Change one diagram's `DiagramBuildEnum` value.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no build has `id`,
    /// and [`super::Error::Refused`] when the record grammar refuses the
    /// mode, leaving the edit unchanged.
    pub fn set_mode(&mut self, id: Id, mode: BuildType) -> Result<()> {
        self.require_build("diagrams.set_mode", id)?;
        Ok(self
            .state_mut()?
            .inner
            .set_mode(id, mode)
            .map_err(scoped::staging("diagrams.set_mode"))?)
    }

    /// Point one diagram's `shapeIdRef` at an existing `OfficeArt` shape of
    /// the slide. Returns the build's new identity.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no build has `id`,
    /// and [`super::Error::Refused`] when the shape does not exist or the
    /// change would duplicate a build identity, leaving the edit unchanged.
    pub fn set_shape_id(&mut self, id: Id, shape_id: u32) -> Result<Id> {
        self.require_build("diagrams.set_shape_id", id)?;
        Ok(self
            .state_mut()?
            .inner
            .set_shape_id(id, shape_id)
            .map_err(scoped::staging("diagrams.set_shape_id"))?)
    }

    /// [`Edit::set_shape_id`] for a contextual [`ShapeRef`].
    ///
    /// # Errors
    ///
    /// As [`Edit::set_shape_id`].
    pub fn set_shape(&mut self, id: Id, shape: ShapeRef) -> Result<Id> {
        self.require_build("diagrams.set_shape", id)?;
        Ok(self
            .state_mut()?
            .inner
            .set_shape(id, shape)
            .map_err(scoped::staging("diagrams.set_shape"))?)
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
                "PPT diagram edit state was not initialized".into(),
            ))
        })?;
        let (target, _) = inner
            .commit()
            .map_err(scoped::publication("diagrams.commit"))?
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
    /// Start an edit of the diagram build metadata of the slide at a
    /// zero-based position.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. Missing slides,
    /// absent build lists, and unreadable records are reported by the first
    /// projection, staging operation, or commit.
    pub fn edit_diagrams(&self, slide: Position) -> Result<Edit> {
        Edit::new(self, slide)
    }

    /// Apply a diagram patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_diagrams(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
