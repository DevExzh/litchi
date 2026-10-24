//! Package-level animation edits (record 0357).
//!
//! [`Package::edit_animations`] stages edits of the persisted animation and
//! timing records of every slide and main master, [`Edit::commit`] publishes a
//! new [`Package`], and [`Package::apply_animations`] replays a [`Patch`] onto
//! the exact package it was made from. Dropping an [`Edit`] discards it.
//! Actions, hyperlinks, sounds, commands and media references stay inert
//! record metadata; the edit never resolves or executes them.

use super::editor::{Editor as RecordEdit, EditorLimits};
use super::{
    AnimationInfo, BuildList, BuildListEntry, Commit, Error, ExtendedTimeNode,
    LegacyShapeAnimation, Patch, Result, Timeline,
};
use crate::Package;
use crate::scoped::{self, Base, Failure, Lazy};

/// A staged animation edit over one immutable [`Package`].
pub struct Edit {
    base: Base,
    inner: Lazy<RecordEdit>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("base", &self.base)
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(package: &Package) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            base,
            inner: Lazy::new(),
        })
    }

    fn initialize_inner(&self) -> std::result::Result<RecordEdit, Failure> {
        // A signed package opens for staging; only a changed commit is
        // refused, since it would invalidate the signature.
        RecordEdit::open(self.base.bytes()?.to_vec(), EditorLimits::default())
            .map_err(scoped::publication("edit_animations"))
    }

    fn inner(&self) -> std::result::Result<&RecordEdit, Failure> {
        self.inner.get_or_try_init(|| self.initialize_inner())
    }

    fn inner_mut(&mut self) -> std::result::Result<&mut RecordEdit, Failure> {
        let base = self.base.clone();
        self.inner.get_or_try_init_mut(|| {
            RecordEdit::open(base.bytes()?.to_vec(), EditorLimits::default())
                .map_err(scoped::publication("edit_animations"))
        })
    }

    fn require_owner(&self, operation: &'static str, persist_id: u32) -> Result<()> {
        if self.inner()?.find(persist_id).is_some() {
            Ok(())
        } else {
            Err(scoped::not_found(operation, "no slide or main master has this persist ID").into())
        }
    }

    fn require_shape(&self, operation: &'static str, persist_id: u32, shape_id: u32) -> Result<()> {
        self.require_owner(operation, persist_id)?;
        if self.inner()?.has_shape(persist_id, shape_id) {
            Ok(())
        } else {
            Err(scoped::not_found(operation, "the slide or main master has no such shape").into())
        }
    }

    /// The staged animation timeline of every slide and main master.
    #[must_use = "handle the fallible animation timelines result"]
    pub fn timelines(&self) -> Result<Vec<Timeline>> {
        Ok(self.inner()?.timelines())
    }

    /// The staged timeline of the slide or main master `persist_id`.
    #[must_use = "handle the fallible animation lookup result"]
    pub fn find(&self, persist_id: u32) -> Result<Option<Timeline>> {
        Ok(self.inner()?.find(persist_id))
    }

    /// The staged legacy `PowerPoint` 97 shape animations.
    #[must_use = "handle the fallible legacy animation result"]
    pub fn legacy_shape_animations(&self) -> Result<Vec<LegacyShapeAnimation>> {
        Ok(self.inner()?.legacy_shape_animations())
    }

    /// The staged legacy animation of shape `shape_id` on `persist_id`.
    #[must_use = "handle the fallible shape animation lookup result"]
    pub fn find_shape(
        &self,
        persist_id: u32,
        shape_id: u32,
    ) -> Result<Option<LegacyShapeAnimation>> {
        Ok(self.inner()?.find_shape(persist_id, shape_id))
    }

    /// Adds a child to the root timing container at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn add(&mut self, persist_id: u32, index: usize, node: ExtendedTimeNode) -> Result<()> {
        self.require_owner("animation.add", persist_id)?;
        Ok(self
            .inner_mut()?
            .add(persist_id, index, node)
            .map_err(scoped::staging("animation.add"))?)
    }

    /// Updates one root child in place.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn update(&mut self, persist_id: u32, index: usize, node: ExtendedTimeNode) -> Result<()> {
        self.require_owner("animation.update", persist_id)?;
        Ok(self
            .inner_mut()?
            .update(persist_id, index, node)
            .map_err(scoped::staging("animation.update"))?)
    }

    /// Replaces the root timeline and build list atomically.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn replace(
        &mut self,
        persist_id: u32,
        root: Option<ExtendedTimeNode>,
        builds: Option<BuildList>,
    ) -> Result<()> {
        self.require_owner("animation.replace", persist_id)?;
        Ok(self
            .inner_mut()?
            .replace(persist_id, root, builds)
            .map_err(scoped::staging("animation.replace"))?)
    }

    /// Removes the root timeline child at `index` and returns it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn remove(&mut self, persist_id: u32, index: usize) -> Result<ExtendedTimeNode> {
        self.require_owner("animation.remove", persist_id)?;
        Ok(self
            .inner_mut()?
            .remove(persist_id, index)
            .map_err(scoped::staging("animation.remove"))?)
    }

    /// Reorders the root timeline children according to `order`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn reorder(&mut self, persist_id: u32, order: &[usize]) -> Result<()> {
        self.require_owner("animation.reorder", persist_id)?;
        Ok(self
            .inner_mut()?
            .reorder(persist_id, order)
            .map_err(scoped::staging("animation.reorder"))?)
    }

    /// Inserts a build-list entry at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn add_build(
        &mut self,
        persist_id: u32,
        index: usize,
        build: BuildListEntry,
    ) -> Result<()> {
        self.require_owner("animation.add_build", persist_id)?;
        Ok(self
            .inner_mut()?
            .add_build(persist_id, index, build)
            .map_err(scoped::staging("animation.add_build"))?)
    }

    /// Replaces the build-list entry at `index`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn update_build(
        &mut self,
        persist_id: u32,
        index: usize,
        build: BuildListEntry,
    ) -> Result<()> {
        self.require_owner("animation.update_build", persist_id)?;
        Ok(self
            .inner_mut()?
            .update_build(persist_id, index, build)
            .map_err(scoped::staging("animation.update_build"))?)
    }

    /// Removes the build-list entry at `index` and returns it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn remove_build(&mut self, persist_id: u32, index: usize) -> Result<BuildListEntry> {
        self.require_owner("animation.remove_build", persist_id)?;
        Ok(self
            .inner_mut()?
            .remove_build(persist_id, index)
            .map_err(scoped::staging("animation.remove_build"))?)
    }

    /// Reorders the build list according to `order`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn reorder_builds(&mut self, persist_id: u32, order: &[usize]) -> Result<()> {
        self.require_owner("animation.reorder_builds", persist_id)?;
        Ok(self
            .inner_mut()?
            .reorder_builds(persist_id, order)
            .map_err(scoped::staging("animation.reorder_builds"))?)
    }

    /// Replaces or clears the legacy `PowerPoint` 97 animation of one shape.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound) when no
    /// slide or main master has `persist_id` or it has no shape
    /// `shape_id`, and with another kind when the record
    /// grammar refuses the request, leaving the edit unchanged.
    pub fn replace_shape_animation(
        &mut self,
        persist_id: u32,
        shape_id: u32,
        animation: Option<AnimationInfo>,
    ) -> Result<()> {
        self.require_shape("animation.replace_shape_animation", persist_id, shape_id)?;
        Ok(self
            .inner_mut()?
            .replace_shape_animation(persist_id, shape_id, animation)
            .map_err(scoped::staging("animation.replace_shape_animation"))?)
    }

    /// Publish the staged animations as a new package. An edit whose staged
    /// records equal the source (nothing staged, or every change reverted)
    /// publishes the source package unchanged; a record rewritten back to
    /// its source animation model keeps its source bytes.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed write to a signed package,
    /// [`Error::Refused`] when the package holds storage the publisher cannot
    /// preserve, and [`Error::Package`] when the result cannot be
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        let Self { base, inner } = self;
        inner
            .get_or_try_init(|| {
                RecordEdit::open(base.bytes()?.to_vec(), EditorLimits::default())
                    .map_err(scoped::publication("edit_animations"))
            })
            .map_err(Error::from)?;
        let inner = inner.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT animation edit state was not initialized".into(),
            ))
        })?;
        let after = match inner.into_net_bytes() {
            Ok(Some(bytes)) => bytes,
            Ok(None) => return Ok(super::published(base.unchanged()?)),
            // The record publisher refuses a changed write of a signed
            // package; report the signature.
            Err(_) if base.is_signed()? => return Err(Error::Signed),
            Err(error) => return Err(scoped::publication("animation.commit")(error).into()),
        };
        Ok(super::published(base.publish(after)?))
    }
}

impl Package {
    /// Start an animation edit over every slide and main master of this
    /// package.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. Persisted-record
    /// and representability errors are returned by the first projection,
    /// staging operation, or commit; a signed package opens and its changed
    /// commit is [`Error::Signed`].
    pub fn edit_animations(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply an animation patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source.
    pub fn apply_animations(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
