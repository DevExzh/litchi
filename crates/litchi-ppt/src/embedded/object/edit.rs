//! Package-level embedded-object edits (record 0357).
//!
//! [`Package::edit_objects`] stages edits of the live document's
//! `ExObjListContainer` OLE definitions and their persisted `ExOleObjStg`
//! storages, [`Edit::commit`] publishes a new [`Package`] through one
//! append-only user edit, and [`Package::apply_objects`] replays a
//! [`Patch`] onto the exact package it was made from. Dropping an [`Edit`]
//! discards it. Storages stay inert: nothing is loaded, activated or run.
//!
//! ```
//! let package = litchi_ppt::Package::new()?;
//! // A presentation without an external-object list has no object to edit.
//! assert!(matches!(
//!     package.edit_objects(),
//!     Err(litchi_ppt::embedded::object::Error::Refused(_))
//! ));
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use litchi_core::edit::{Refusal, RefusalKind};

use super::{Collection, Commit, Editor, Error, ExternalObject, Patch, Result};
use crate::Package;
use crate::embedded::storage::Storage;
use crate::package::Error as PackageError;
use crate::scoped::{self, Base, Failure};

const MISSING_OBJECT: &str = "no embedded object has this id";
const MISSING_STORAGE: &str = "no embedded object references this persist id";

/// A staged edit of the live document's embedded OLE objects over one
/// immutable [`Package`].
pub struct Edit {
    base: Base,
    source: Collection,
    inner: Editor,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("objects", &self.inner.objects().objects.len())
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
        let source = package.presentation()?.ole_objects()?.ok_or_else(|| {
            scoped::unsupported_target(
                "embedded_object.edit",
                "the presentation has no external-object list",
            )
        })?;
        let inner = match Editor::open(base.bytes().to_vec(), source.clone()) {
            Ok(inner) => inner,
            // The record publisher refuses the signature storages any write
            // would invalidate; report the signature, not the storage.
            Err(_) if base.is_signed()? => return Err(Error::Signed),
            Err(error) => return Err(scoped::publication("embedded_object.edit")(error).into()),
        };
        Ok(Self {
            base,
            source,
            inner,
        })
    }

    fn require_object(&self, operation: &'static str, id: u32) -> Result<()> {
        if self.inner.objects().get(id).is_some() {
            Ok(())
        } else {
            Err(scoped::not_found(operation, MISSING_OBJECT).into())
        }
    }

    /// The staged embedded-object collection.
    #[must_use]
    pub fn objects(&self) -> &Collection {
        self.inner.objects()
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
        let seed = self.inner.collection.id_seed;
        self.inner.collection.id_seed = seed.max(object.id());
        self.inner.add(object, storage).map_err(|error| {
            self.inner.collection.id_seed = seed;
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
            .inner
            .reorder(ids)
            .map_err(staging("embedded_object.reorder"))?)
    }

    /// Publish the staged objects as a new package. An edit with no net
    /// change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the package holds content the
    /// publisher cannot preserve and [`Error::Package`] when the result
    /// cannot be republished.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            base,
            source,
            inner,
        } = self;
        if !inner
            .has_net_change(&source)
            .map_err(scoped::publication("embedded_object.commit"))?
        {
            return Ok(super::published(base.unchanged()));
        }
        let after = inner
            .finish()
            .map_err(scoped::publication("embedded_object.commit"))?;
        Ok(super::published(base.publish(after)?))
    }
}

impl Package {
    /// Start an edit of the live document's embedded OLE objects.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Package`] for an encrypted presentation or one whose
    /// object list cannot be read, [`Error::Signed`] for a signed one, and
    /// [`Error::Refused`] with [`RefusalKind::UnsupportedTarget`] when the
    /// live document has no external-object list.
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
