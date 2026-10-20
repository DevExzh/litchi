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
    inner: Editor,
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
        let bytes = base.to_vec();
        let options = base.options();
        let inner = Editor::open_with_options(bytes, options)?;
        Ok(Self { base, inner })
    }

    /// The current immutable semantic envelope snapshot.
    #[must_use]
    pub fn envelope(&self) -> &TransactionSnapshot {
        self.inner.envelope()
    }

    /// The current decoded envelope, when present.
    #[must_use]
    pub fn value(&self) -> Option<&Envelope> {
        self.inner.value()
    }

    /// The current supported Office message body, when present.
    #[must_use]
    pub fn message(&self) -> Option<&Message> {
        self.inner.message()
    }

    /// Set or replace the complete envelope FIB range.
    pub fn set(&mut self, envelope: Envelope) -> Result<()> {
        self.inner.set(envelope).map(|_| ()).map_err(Error::from)
    }

    /// Set or replace the supported Office message body.
    pub fn set_message(&mut self, message: Message) -> Result<()> {
        self.inner
            .set_message(message)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Replace an existing envelope, rejecting a missing FIB range.
    pub fn replace(&mut self, envelope: Envelope) -> Result<()> {
        self.inner
            .replace(envelope)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Apply and publish a clone-first semantic transaction atomically.
    pub fn apply(&mut self, transaction: RecordEdit) -> Result<()> {
        self.inner
            .apply(transaction)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Clone-first update of the supported Office message body.
    pub fn update_message<F>(&mut self, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Message),
    {
        self.inner
            .update_message(edit)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Clear the complete envelope FIB range.
    pub fn clear(&mut self) -> Result<()> {
        self.inner.clear().map(|_| ()).map_err(Error::from)
    }

    /// Validate the staged edits and publish them as a new package.
    ///
    /// # Errors
    ///
    /// Returns an error when the edited part fails validation, the package
    /// cannot be republished, or [`Error::Signed`] for a changed signed
    /// package.
    pub fn commit(self) -> Result<Commit> {
        let after = self.inner.finish()?;
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
