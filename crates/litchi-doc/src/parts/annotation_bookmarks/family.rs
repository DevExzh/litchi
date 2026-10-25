//! The annotation bookmark edit family (record 0351 §2.3, record 0359).
//!
//! [`crate::Package::edit_annotation_bookmarks`] stages annotation bookmark edits on an [`Edit`];
//! [`Edit::commit`] publishes a new [`crate::Package`] with an exact
//! whole-package [`Patch`] and [`Diagnostics`];
//! [`crate::Package::apply_annotation_bookmarks`] replays the patch. The part-level codec
//! stays public in [`super::record`]; the byte-level editor is crate-private.

#![allow(
    unused_imports,
    reason = "the family forwards the part editor's signatures verbatim"
)]

use super::codec;
use super::model::{Tag, Tags};
use super::package::Editor;
use super::transaction::{
    Edit as RecordEdit, Error as TransactionError, Patch as RecordPatch,
    Snapshot as TransactionSnapshot,
};
use super::validation;
use crate::package::Error as PackageError;
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, PackagePatch, ProtectionPolicy, classify};
use litchi_ole_common::object::{Edit as ObjectEditor, Limits, Targets};

crate::scoped::family!("annotation bookmark");

/// A staged annotation bookmark edit of one package. Dropping it discards it.
#[derive(Clone)]
pub struct Edit {
    base: crate::scoped::Base,
    inner: Option<Editor>,
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
    fn open(base: crate::scoped::Base) -> Result<Self> {
        Ok(Self { base, inner: None })
    }

    fn inner_mut(&mut self) -> Result<&mut Editor> {
        if self.inner.is_none() {
            let bytes = self.base.to_vec()?;
            let options = self.base.options();
            self.inner = Some(Editor::open_with_options(
                bytes,
                self.base.source_limits(),
                options,
            )?);
        }
        self.inner.as_mut().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "annotation bookmark edit was not initialized".into(),
            ))
        })
    }

    fn take_inner(&mut self) -> Result<Editor> {
        self.inner_mut()?;
        self.inner.take().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "annotation bookmark edit was not initialized".into(),
            ))
        })
    }

    /// Current immutable semantic table state.
    pub fn tags(&mut self) -> Result<&TransactionSnapshot> {
        Ok(self.inner_mut()?.tags())
    }

    /// Decoded tags, when present.
    pub fn value(&mut self) -> Result<Option<&Tags>> {
        Ok(self.inner_mut()?.value())
    }

    /// Set or replace the complete present table.
    pub fn set(&mut self, tags: Tags) -> Result<()> {
        self.inner_mut()?.set(tags).map(|_| ()).map_err(Error::from)
    }

    /// Replace the table, requiring an existing FIB range.
    pub fn replace(&mut self, tags: Tags) -> Result<()> {
        self.inner_mut()?
            .replace(tags)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Apply and publish a semantic transaction atomically.
    pub fn apply(&mut self, transaction: RecordEdit) -> Result<()> {
        self.inner_mut()?
            .apply(transaction)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Replace one annotation tag and publish atomically.
    pub fn replace_entry(&mut self, index: usize, tag: Tag) -> Result<()> {
        self.inner_mut()?
            .replace_entry(index, tag)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Insert one annotation tag and publish atomically.
    pub fn insert(&mut self, index: usize, tag: Tag) -> Result<()> {
        self.inner_mut()?
            .insert(index, tag)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Remove one annotation tag and publish atomically.
    pub fn remove(&mut self, index: usize) -> Result<()> {
        self.inner_mut()?
            .remove(index)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Remove the complete FIB range.
    pub fn clear(&mut self) -> Result<()> {
        self.inner_mut()?.clear().map(|_| ()).map_err(Error::from)
    }

    /// Validate the staged edits and publish them as a new package.
    ///
    /// # Errors
    ///
    /// Returns an error when the edited part fails validation, the package
    /// cannot be republished, or [`Error::Signed`] for a changed signed
    /// package.
    pub fn commit(mut self) -> Result<Commit> {
        let after = self.take_inner()?.finish()?;
        Ok(published(self.base.publish(after)?))
    }
}

impl crate::Package {
    /// Start a annotation bookmark edit of this package under its retained edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's annotation bookmark cannot be read for editing.
    pub fn edit_annotation_bookmarks(&self) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture(self)?)
    }

    /// Start a annotation bookmark edit of this package under explicit edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's annotation bookmark cannot be read for editing.
    pub fn edit_annotation_bookmarks_with_options(
        &self,
        options: crate::EditOptions,
    ) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture_with_options(self, options)?)
    }

    /// Replay a annotation bookmark patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed package.
    pub fn apply_annotation_bookmarks(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }
}
