//! Package-level picture-bullet edits (record 0357).
//!
//! [`Package::edit_picture_bullets`] stages lifecycle edits of the live
//! document's picture-bullet collection, [`Edit::commit`] publishes a new
//! [`Package`] through one append-only user edit, and
//! [`Package::apply_picture_bullets`] replays a [`Patch`] onto the exact
//! package it was made from. Dropping an [`Edit`] discards it.

use litchi_core::edit::{Refusal, RefusalKind};

use super::record::{PictureBulletPackageLimits, PictureBulletSnapshot, RecordEdit};
use super::{
    Commit, Error, Patch, PictureBullet, PictureBulletChange, PictureBulletCollection, Result,
};
use crate::Package;
use crate::scoped::{self, Base, Failure, Lazy};

const ABSENT_COLLECTION: &str = "the live document has no picture-bullet collection";

/// A staged picture-bullet edit over one immutable [`Package`].
#[derive(Debug)]
pub struct Edit {
    base: Base,
    inner: Lazy<RecordEdit>,
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
        let snapshot = PictureBulletSnapshot::from_arc(
            self.base.shared_bytes()?,
            PictureBulletPackageLimits::default(),
        )
        .map_err(scoped::publication("edit_picture_bullets"))?;
        snapshot
            .edit()
            .map_err(scoped::publication("edit_picture_bullets"))
    }

    fn inner(&self) -> std::result::Result<&RecordEdit, Failure> {
        self.inner.get_or_try_init(|| self.initialize_inner())
    }

    fn inner_mut(&mut self) -> std::result::Result<&mut RecordEdit, Failure> {
        let base = self.base.clone();
        self.inner.get_or_try_init_mut(|| {
            let snapshot = PictureBulletSnapshot::from_arc(
                base.shared_bytes()?,
                PictureBulletPackageLimits::default(),
            )
            .map_err(scoped::publication("edit_picture_bullets"))?;
            snapshot
                .edit()
                .map_err(scoped::publication("edit_picture_bullets"))
        })
    }

    /// The staged picture-bullet collection, or `None` when the live
    /// document has none.
    #[must_use = "handle the fallible picture-bullets projection result"]
    pub fn picture_bullets(&self) -> Result<Option<&PictureBulletCollection>> {
        Ok(self.inner()?.picture_bullets())
    }

    /// The staged operations, in order.
    #[must_use = "handle the fallible picture-bullets changes result"]
    pub fn changes(&self) -> Result<&[PictureBulletChange]> {
        Ok(self.inner()?.changes())
    }

    fn require_index(&self, operation: &'static str, index: u16) -> Result<()> {
        let Some(collection) = self.inner()?.picture_bullets() else {
            return Err(scoped::unsupported_target(operation, ABSENT_COLLECTION).into());
        };
        if collection
            .bullets
            .iter()
            .any(|bullet| bullet.index == index)
        {
            Ok(())
        } else {
            Err(scoped::not_found(operation, "no picture bullet has this index").into())
        }
    }

    /// Stage a replacement of the bullet at `index`; the replacement keeps
    /// that index. An equal bullet is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection or index is absent, or
    /// the bullet changes its index or cannot be stored, leaving the edit
    /// unchanged.
    pub fn replace_picture_bullet(&mut self, index: u16, bullet: PictureBullet) -> Result<()> {
        self.require_index("picture_bullets.replace_picture_bullet", index)?;
        Ok(self
            .inner_mut()?
            .replace_picture_bullet(index, bullet)
            .map_err(scoped::staging("picture_bullets.replace_picture_bullet"))?)
    }

    /// Stage an append of `bullet` at its explicit, currently unused index,
    /// returning that index.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection is absent, the index is
    /// already used, or the bullet cannot be stored.
    pub fn append_picture_bullet(&mut self, bullet: PictureBullet) -> Result<u16> {
        if self.inner()?.picture_bullets().is_none() {
            return Err(scoped::unsupported_target(
                "picture_bullets.append_picture_bullet",
                ABSENT_COLLECTION,
            )
            .into());
        }
        Ok(self
            .inner_mut()?
            .append_picture_bullet(bullet)
            .map_err(scoped::staging("picture_bullets.append_picture_bullet"))?)
    }

    /// Stage removal of the bullet at `index`, returning it. Removal is
    /// admitted only when every discoverable `TextPFException9` user proves
    /// the bullet is unreferenced.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the collection or index is absent, or
    /// the bullet is referenced, leaving the edit unchanged.
    pub fn remove_picture_bullet(&mut self, index: u16) -> Result<PictureBullet> {
        self.require_index("picture_bullets.remove_picture_bullet", index)?;
        Ok(self
            .inner_mut()?
            .remove_picture_bullet(index)
            .map_err(scoped::staging("picture_bullets.remove_picture_bullet"))?)
    }

    /// Re-submitting the current order is an exact no-op; any other order is
    /// refused, because it would require an exhaustive reference rewrite
    /// across all shape and master owners.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with
    /// [`RefusalKind::UnsupportedOperation`](litchi_core::edit::RefusalKind::UnsupportedOperation)
    /// for a changed order, or
    /// [`RefusalKind::UnsupportedTarget`](litchi_core::edit::RefusalKind::UnsupportedTarget)
    /// when the collection is absent and `order` is not empty.
    pub fn reorder_picture_bullets(&mut self, order: &[u16]) -> Result<()> {
        const OPERATION: &str = "picture_bullets.reorder_picture_bullets";
        match self.inner_mut()?.reorder_picture_bullets(order) {
            Ok(()) => Ok(()),
            Err(_) if self.inner()?.picture_bullets().is_none() => {
                Err(scoped::unsupported_target(OPERATION, ABSENT_COLLECTION).into())
            },
            Err(_) => Err(Error::Refused(Refusal::new(
                RefusalKind::UnsupportedOperation,
                OPERATION,
                "picture-bullet reordering requires a complete TextPFException9 reference remap",
            ))),
        }
    }

    /// Publish the staged picture bullets as a new package. An edit with no
    /// net change publishes the source package unchanged.
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
                let snapshot = PictureBulletSnapshot::from_arc(
                    base.shared_bytes()?,
                    PictureBulletPackageLimits::default(),
                )
                .map_err(scoped::publication("edit_picture_bullets"))?;
                snapshot
                    .edit()
                    .map_err(scoped::publication("edit_picture_bullets"))
            })
            .map_err(Error::from)?;
        let inner = inner.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT picture-bullet edit state was not initialized".into(),
            ))
        })?;
        let after = match inner.commit() {
            Ok(commit) => commit.snapshot().bytes().to_vec(),
            // The publisher refuses the signature storages a changed write
            // would invalidate; report the signature, not the storage.
            Err(_) if base.is_signed()? => return Err(Error::Signed),
            Err(error) => {
                return Err(scoped::publication("picture_bullets.commit")(error).into());
            },
        };
        Ok(super::published(base.publish(after)?))
    }
}

impl Package {
    /// Start a picture-bullet edit over this package.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. A live document
    /// owner that cannot be read or losslessly represented is reported by the
    /// first projection, staging operation, or commit.
    pub fn edit_picture_bullets(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply a picture-bullet patch to the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source.
    pub fn apply_picture_bullets(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
