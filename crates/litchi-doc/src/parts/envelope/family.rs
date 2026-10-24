//! The envelope edit family (record 0351 §2.3, record 0359).
//!
//! [`crate::Package::edit_envelope`] stages envelope edits on an [`Edit`];
//! [`Edit::commit`] publishes a new [`crate::Package`] with an exact
//! whole-package [`Patch`] and [`Diagnostics`];
//! [`crate::Package::apply_envelope`] replays the patch. The part-level codec
//! stays public in [`super::record`]; the byte-level editor is crate-private.

#![allow(
    unused_imports,
    reason = "the family forwards the part editor's signatures verbatim"
)]

use super::codec;
use super::model::{Envelope, Message};
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

crate::scoped::family!("envelope");

/// A staged envelope edit of one package. Dropping it discards it.
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
                "envelope edit was not initialized".into(),
            ))
        })
    }

    fn take_inner(&mut self) -> Result<Editor> {
        self.inner_mut()?;
        self.inner.take().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "envelope edit was not initialized".into(),
            ))
        })
    }

    /// The current immutable semantic envelope snapshot.
    pub fn envelope(&mut self) -> Result<&TransactionSnapshot> {
        Ok(self.inner_mut()?.envelope())
    }

    /// The current decoded envelope, when present.
    pub fn value(&mut self) -> Result<Option<&Envelope>> {
        Ok(self.inner_mut()?.value())
    }

    /// The current supported Office message body, when present.
    pub fn message(&mut self) -> Result<Option<&Message>> {
        Ok(self.inner_mut()?.message())
    }

    /// Set or replace the complete envelope FIB range.
    pub fn set(&mut self, envelope: Envelope) -> Result<()> {
        self.inner_mut()?
            .set(envelope)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Set or replace the supported Office message body.
    pub fn set_message(&mut self, message: Message) -> Result<()> {
        self.inner_mut()?
            .set_message(message)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Replace an existing envelope, rejecting a missing FIB range.
    pub fn replace(&mut self, envelope: Envelope) -> Result<()> {
        self.inner_mut()?
            .replace(envelope)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Apply and publish a clone-first semantic transaction atomically.
    pub fn apply(&mut self, transaction: RecordEdit) -> Result<()> {
        self.inner_mut()?
            .apply(transaction)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Clone-first update of the supported Office message body.
    pub fn update_message<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Message),
    {
        self.inner_mut()?
            .update_message(edit)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Clear the complete envelope FIB range.
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
    /// Start a envelope edit of this package under its retained edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's envelope cannot be read for editing.
    pub fn edit_envelope(&self) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture(self)?)
    }

    /// Start a envelope edit of this package under explicit edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's envelope cannot be read for editing.
    pub fn edit_envelope_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture_with_options(self, options)?)
    }

    /// Replay a envelope patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed package.
    pub fn apply_envelope(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }
}
