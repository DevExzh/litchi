//! Package-level embedded-object edits (record 0357).
//!
//! [`Package::edit_objects`] stages edits of the live document's
//! `ExObjListContainer` OLE definitions and their persisted `ExOleObjStg`
//! storages, [`Edit::commit`] publishes a new [`Package`] through one
//! append-only user edit, and [`Package::apply_objects`] replays a
//! [`Patch`] onto the exact package it was made from. Dropping an [`Edit`]
//! discards it. Storages stay inert: nothing is loaded, activated or run.
//!
//! A presentation without an external-object list starts from an empty one;
//! the first committed object creates the `ExObjListContainer` (with its
//! `ExObjListAtom` identifier seed) in the live document.
//!
//! ```
//! let package = litchi_ppt::Package::new()?;
//! let edit = package.edit_objects()?;
//! assert!(edit.objects()?.objects.is_empty());
//! // Nothing staged: the commit is the source package unchanged.
//! assert!(!edit.commit()?.diagnostics().changed());
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use litchi_core::edit::{Refusal, RefusalKind};

use super::{Collection, Commit, Editor, Error, ExternalObject, Patch, Result};
use crate::Package;
use crate::embedded::storage::Storage;
use crate::package::Error as PackageError;
use crate::scoped::{self, Base, Failure, Lazy};

const MISSING_OBJECT: &str = "no embedded object has this id";
const MISSING_STORAGE: &str = "no embedded object references this persist id";

/// A staged edit of the live document's embedded OLE objects over one
/// immutable [`Package`].
pub struct Edit {
    base: Base,
    state: Lazy<State>,
}

struct State {
    source: Collection,
    inner: Editor,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("initialized", &self.state.get().is_some())
            .finish_non_exhaustive()
    }
}

/// Map a staging failure: a request the object-list grammar refuses becomes
/// a content-free [`RefusalKind::InvalidValue`] refusal.
fn staging(operation: &'static str) -> impl FnOnce(PackageError) -> Failure {
    move |error| match error {
        PackageError::Corrupted(_) | PackageError::InvalidFormat(_) => {
            Failure::Refused(Refusal::new(
                RefusalKind::InvalidValue,
                operation,
                "the external-object list cannot hold the requested change",
            ))
        },
        error => Failure::Package(error),
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
        // A document without an `ExObjListContainer` edits an empty list; a
        // committed change inserts the container ([MS-PPT] 2.4.1).
        let presentation = self.base.package().presentation().map_err(Failure::from)?;
        let source_limits = presentation.source_limits();
        let source = presentation
            .ole_objects()?
            .unwrap_or_else(Collection::empty);
        // A signed package opens for staging; only a changed commit is
        // refused, since it would invalidate the signature.
        let inner = Editor::open_admitting_signature(
            self.base.bytes()?.to_vec(),
            source.clone(),
            self.base.cfb_limits()?,
            source_limits,
        )
        .map_err(scoped::publication("embedded_object.edit"))?;
        Ok(State { source, inner })
    }

    fn state(&self) -> std::result::Result<&State, Failure> {
        self.state.get_or_try_init(|| self.initialize_state())
    }

    fn state_mut(&mut self) -> std::result::Result<&mut State, Failure> {
        let base = self.base.clone();
        self.state.get_or_try_init_mut(|| {
            let presentation = base.package().presentation().map_err(Failure::from)?;
            let source_limits = presentation.source_limits();
            let source = presentation
                .ole_objects()?
                .unwrap_or_else(Collection::empty);
            let inner = Editor::open_admitting_signature(
                base.bytes()?.to_vec(),
                source.clone(),
                base.cfb_limits()?,
                source_limits,
            )
            .map_err(scoped::publication("embedded_object.edit"))?;
            Ok::<State, Failure>(State { source, inner })
        })
    }

    fn require_object(&self, operation: &'static str, id: u32) -> Result<()> {
        if self.state()?.inner.objects().get(id).is_some() {
            Ok(())
        } else {
            Err(scoped::not_found(operation, MISSING_OBJECT).into())
        }
    }

    /// The staged embedded-object collection.
    #[must_use = "handle the fallible embedded-object projection result"]
    pub fn objects(&self) -> Result<&Collection> {
        Ok(self.state()?.inner.objects())
    }

    /// Stage a new external object and its persisted storage record,
    /// returning the persist id allocated for the storage (the object's own
    /// persist id is replaced by it). An object id above the list's
    /// identifier seed raises the seed to that id, as `[MS-PPT]`
    /// `ExObjListAtom` requires the seed to cover every object id.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the object's id is zero, already used
    /// or beyond the signed seed domain, or the storage cannot be stored,
    /// leaving the edit unchanged.
    pub fn add(&mut self, object: ExternalObject, storage: Storage) -> Result<u32> {
        let state = self.state_mut()?;
        let seed = state.inner.collection.id_seed;
        state.inner.collection.id_seed = seed.max(object.id());
        state.inner.add(object, storage).map_err(|error| {
            state.inner.collection.id_seed = seed;
            Error::from(staging("embedded_object.add")(error))
        })
    }

    /// Stage a replacement of the storage record behind `persist_id`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`] when no
    /// object references `persist_id`, and [`Error::Refused`] when the
    /// storage cannot be stored, leaving the edit unchanged.
    pub fn replace_storage(&mut self, persist_id: u32, storage: Storage) -> Result<()> {
        if !self
            .state()?
            .inner
            .objects()
            .objects
            .iter()
            .any(|object| object.persist_id() == persist_id)
        {
            return Err(
                scoped::not_found("embedded_object.replace_storage", MISSING_STORAGE).into(),
            );
        }
        Ok(self
            .state_mut()?
            .inner
            .replace_storage(persist_id, storage)
            .map_err(staging("embedded_object.replace_storage"))?)
    }

    /// Stage removal of the object with `id` and, when no other object
    /// references it, its persisted storage. Returns the removed object.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`] when no
    /// object has `id`, leaving the edit unchanged.
    pub fn remove(&mut self, id: u32) -> Result<ExternalObject> {
        self.require_object("embedded_object.remove", id)?;
        Ok(self
            .state_mut()?
            .inner
            .remove(id)
            .map_err(staging("embedded_object.remove"))?)
    }

    /// Stage a new order of the external-object list; `ids` names every
    /// object exactly once.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with [`RefusalKind::NotFound`] when an id
    /// names no object, and [`Error::Refused`] when `ids` repeats or omits an
    /// object, leaving the edit unchanged.
    pub fn reorder(&mut self, ids: &[u32]) -> Result<()> {
        for id in ids {
            self.require_object("embedded_object.reorder", *id)?;
        }
        Ok(self
            .state_mut()?
            .inner
            .reorder(ids)
            .map_err(staging("embedded_object.reorder"))?)
    }

    /// Publish the staged objects as a new package. An edit with no net
    /// change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed write to a signed package,
    /// [`Error::Refused`] when the package holds content the publisher
    /// cannot preserve and [`Error::Package`] when the result cannot be
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        let Self { base, state } = self;
        state
            .get_or_try_init(|| {
                let presentation = base.package().presentation().map_err(Failure::from)?;
                let source_limits = presentation.source_limits();
                let source = presentation
                    .ole_objects()?
                    .unwrap_or_else(Collection::empty);
                let inner = Editor::open_admitting_signature(
                    base.bytes()?.to_vec(),
                    source.clone(),
                    base.cfb_limits()?,
                    source_limits,
                )
                .map_err(scoped::publication("embedded_object.edit"))?;
                Ok::<State, Failure>(State { source, inner })
            })
            .map_err(Error::from)?;
        let State { source, inner } = state.into_inner().ok_or_else(|| {
            Failure::Package(PackageError::Corrupted(
                "PPT embedded-object edit state was not initialized".into(),
            ))
        })?;
        if !inner
            .has_net_change(&source)
            .map_err(scoped::publication("embedded_object.commit"))?
        {
            return Ok(super::published(base.unchanged()?));
        }
        let after = match inner.finish() {
            Ok(bytes) => bytes,
            // The record publisher refuses a changed write of a signed
            // package; report the signature.
            Err(_) if base.is_signed()? => return Err(Error::Signed),
            Err(error) => return Err(scoped::publication("embedded_object.commit")(error).into()),
        };
        Ok(super::published(base.publish(after)?))
    }
}

impl Package {
    /// Start an edit of the live document's embedded OLE objects.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. An object list that
    /// cannot be read is reported by the first projection, staging operation,
    /// or commit. A presentation without an external-object list edits an
    /// empty one; a signed presentation opens, and its changed commit is
    /// [`Error::Signed`].
    pub fn edit_objects(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply an embedded-object patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source.
    pub fn apply_objects(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
