//! Package-level external-media edits (record 0357).
//!
//! [`Package::edit_external_media`] stages edits of the live document's
//! `ExObjListContainer` media definitions, [`Edit::commit`] publishes a new
//! [`Package`] through one append-only user edit that rewrites only the live
//! document, and [`Package::apply_external_media`] replays a [`Patch`] onto
//! the exact package it was made from. Dropping an [`Edit`] discards it.
//! Paths stay inert metadata: nothing is resolved, opened or played.

use super::transaction::{RecordEdit, Snapshot};
use super::{Change, Collection, Commit, Object, Patch, Playback, Result};
use crate::Package;
use crate::scoped::{self, Base, Failure, Lazy};
use crate::slide_order::{self, records::Owner};

const MISSING_OBJECT: &str = "no external media object has this id";

/// A staged edit of the live document's external-media definitions over one
/// immutable [`Package`].
pub struct Edit {
    base: Base,
    state: Lazy<State>,
}

struct State {
    native: slide_order::Snapshot,
    source: Snapshot,
    inner: RecordEdit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("initialized", &self.state.get().is_some())
            .finish_non_exhaustive()
    }
}

impl Edit {
    fn new(package: &Package) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            base,
            state: Lazy::new(),
        })
    }

    fn initialize_state(&self) -> std::result::Result<State, Failure> {
        let native = self.base.native()?;
        let source = slide_order::external_media_snapshot(&native).map_err(Failure::from)?;
        let inner = source.edit();
        Ok(State {
            native,
            source,
            inner,
        })
    }

    fn state(&self) -> std::result::Result<&State, Failure> {
        self.state.get_or_try_init(|| self.initialize_state())
    }

    fn state_mut(&mut self) -> std::result::Result<&mut State, Failure> {
        let base = self.base.clone();
        self.state.get_or_try_init_mut(|| {
            let native = base.native()?;
            let source = slide_order::external_media_snapshot(&native).map_err(Failure::from)?;
            let inner = source.edit();
            Ok(State {
                native,
                source,
                inner,
            })
        })
    }

    fn require_object(&self, operation: &'static str, id: u32) -> Result<()> {
        if self
            .state()?
            .inner
            .objects()
            .iter()
            .any(|object| object.id() == id)
        {
            Ok(())
        } else {
            Err(scoped::not_found(operation, MISSING_OBJECT).into())
        }
    }

    /// The staged media collection, or `None` when the live document has no
    /// external-object list.
    #[must_use = "handle the fallible external-media collection result"]
    pub fn collection(&self) -> Result<Option<&Collection>> {
        Ok(self.state()?.inner.collection())
    }

    /// The staged media objects, or an empty slice when no list exists.
    #[must_use = "handle the fallible external-media objects result"]
    pub fn objects(&self) -> Result<&[Object]> {
        Ok(self.state()?.inner.objects())
    }

    /// The staged operations, in order.
    #[must_use = "handle the fallible external-media changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.state()?.inner.changes())
    }

    /// Number of shapes in the source whose `ExObjRefAtom` references the
    /// media object `id`.
    #[must_use = "handle the fallible external-media owner count result"]
    pub fn owner_count(&self, id: u32) -> Result<usize> {
        Ok(self.state()?.source.owner_count(id))
    }

    /// Change or clear the inert path of a movie or linked-audio object.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no object has `id`,
    /// and [`super::Error::Refused`] when the object carries no path or the
    /// record grammar refuses the path, leaving the edit unchanged.
    pub fn set_path(&mut self, id: u32, path: Option<String>) -> Result<()> {
        self.require_object("external_media.set_path", id)?;
        Ok(self
            .state_mut()?
            .inner
            .set_path(id, path)
            .map_err(scoped::staging("external_media.set_path"))?)
    }

    /// Change the inert playback flags of one media object.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no object has `id`,
    /// and [`super::Error::Refused`] when the record grammar refuses the
    /// flags, leaving the edit unchanged.
    pub fn set_playback(&mut self, id: u32, playback: Playback) -> Result<()> {
        self.require_object("external_media.set_playback", id)?;
        Ok(self
            .state_mut()?
            .inner
            .set_playback(id, playback)
            .map_err(scoped::staging("external_media.set_playback"))?)
    }

    /// Insert a typed media definition, allocating no external content. A
    /// live document without an external-object list gains one.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] when the id is zero, already used by
    /// a media object, hyperlink or owner, or the record grammar refuses the
    /// object, leaving the edit unchanged.
    pub fn insert(&mut self, object: Object) -> Result<()> {
        Ok(self
            .state_mut()?
            .inner
            .insert(object)
            .map_err(scoped::staging("external_media.insert"))?)
    }

    /// Replace the definition with `id`; the replacement keeps that id.
    /// Returns the replaced definition.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no object has `id`,
    /// and [`super::Error::Refused`] when the replacement changes its id or
    /// cannot be stored, leaving the edit unchanged.
    pub fn replace(&mut self, id: u32, replacement: Object) -> Result<Object> {
        self.require_object("external_media.replace", id)?;
        Ok(self
            .state_mut()?
            .inner
            .replace(id, replacement)
            .map_err(scoped::staging("external_media.replace"))?)
    }

    /// Remove the definition with `id`. Returns the removed definition.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::Refused`] with
    /// [`litchi_core::edit::RefusalKind::NotFound`] when no object has `id`,
    /// and with [`litchi_core::edit::RefusalKind::UnsupportedTarget`] when a
    /// shape still references it, leaving the edit unchanged.
    pub fn remove(&mut self, id: u32) -> Result<Object> {
        self.require_object("external_media.remove", id)?;
        if self.state()?.source.owner_count(id) > 0 {
            return Err(scoped::unsupported_target(
                "external_media.remove",
                "a shape still references this external media object",
            )
            .into());
        }
        Ok(self
            .state_mut()?
            .inner
            .remove(id)
            .map_err(scoped::staging("external_media.remove"))?)
    }

    /// Publish the staged media definitions as a new package. An edit with
    /// no net change publishes the source package unchanged.
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
        let Self { base, state } = self;
        let State {
            native,
            source,
            inner,
        } = state.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT external-media edit state was not initialized".into(),
            ))
        })?;
        let (target, _) = inner
            .commit()
            .map_err(scoped::publication("external_media.commit"))?
            .into_parts();
        Ok(super::published(base.publish_owner(
            &native,
            Owner::Document,
            source.bytes().to_vec(),
            target.bytes().to_vec(),
        )?))
    }
}

impl Package {
    /// Start an edit of the live document's external-media definitions.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. An unreadable
    /// presentation is reported by the first projection, staging operation,
    /// or commit.
    pub fn edit_external_media(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply an external-media patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::PatchConflict`] when this package is not the
    /// patch's source.
    pub fn apply_external_media(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
