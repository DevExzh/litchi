//! The captions edit family (record 0351 §2.3, record 0359).
//!
//! [`crate::Package::edit_captions`] stages captions edits on an [`Edit`];
//! [`Edit::commit`] publishes a new [`crate::Package`] with an exact
//! whole-package [`Patch`] and [`Diagnostics`];
//! [`crate::Package::apply_captions`] replays the patch. The part-level codec
//! stays public in [`super::record`]; the byte-level editor is crate-private.

#![allow(
    unused_imports,
    reason = "the family forwards the part editor's signatures verbatim"
)]

use super::codec::{AUTO_CAPTION_FIB_INDEX, CAPTION_FIB_INDEX};
use super::package::Editor;
use super::semantic::Tables;
use super::transaction::{
    Edit as RecordEdit, Error as TransactionError, Patch as RecordPatch,
    Snapshot as TransactionSnapshot,
};
use super::validation;
use super::{AutoTable, LabelTable};
use crate::package::Error as PackageError;
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, PackagePatch, ProtectionPolicy, classify};
use litchi_ole_common::object::{Edit as ObjectEditor, Limits, Targets};

crate::scoped::family!("captions");

/// A staged captions edit of one package. Dropping it discards it.
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
            self.inner = Some(Editor::open_with_options(bytes, options)?);
        }
        self.inner.as_mut().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "captions edit was not initialized".into(),
            ))
        })
    }

    fn take_inner(&mut self) -> Result<Editor> {
        self.inner_mut()?;
        self.inner.take().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "captions edit was not initialized".into(),
            ))
        })
    }

    /// Returns the current immutable caption snapshot.
    pub fn captions(&mut self) -> Result<&TransactionSnapshot> {
        Ok(self.inner_mut()?.captions())
    }

    /// Returns the current paired caption tables.
    pub fn tables(&mut self) -> Result<&Tables> {
        Ok(self.inner_mut()?.tables())
    }

    /// Sets or replaces both caption tables atomically.
    pub fn set(&mut self, tables: Tables) -> Result<()> {
        self.inner_mut()?
            .set(tables)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Replaces existing caption metadata, rejecting an absent pair.
    pub fn replace(&mut self, tables: Tables) -> Result<()> {
        self.inner_mut()?
            .replace(tables)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Applies a semantic transaction and publishes it atomically.
    pub fn apply(&mut self, transaction: RecordEdit) -> Result<()> {
        self.inner_mut()?
            .apply(transaction)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Replaces the label table while retaining automatic-caption rules.
    pub fn replace_labels(&mut self, labels: LabelTable) -> Result<()> {
        self.inner_mut()?
            .replace_labels(labels)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Replaces automatic-caption rules while retaining labels.
    pub fn replace_auto(&mut self, auto: AutoTable) -> Result<()> {
        self.inner_mut()?
            .replace_auto(auto)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Removes only the label range, rejecting rules that would dangle.
    pub fn clear_labels(&mut self) -> Result<()> {
        self.inner_mut()?
            .clear_labels()
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Removes only the automatic-caption range while retaining labels.
    pub fn clear_auto(&mut self) -> Result<()> {
        self.inner_mut()?
            .clear_auto()
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Removes both caption table ranges while preserving their old bytes.
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
    /// Start a captions edit of this package under its retained edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's captions cannot be read for editing.
    pub fn edit_captions(&self) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture(self)?)
    }

    /// Start a captions edit of this package under explicit edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's captions cannot be read for editing.
    pub fn edit_captions_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture_with_options(self, options)?)
    }

    /// Replay a captions patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed package.
    pub fn apply_captions(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }
}
