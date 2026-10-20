//! The route slip edit family (record 0351 §2.3, record 0359).
//!
//! [`crate::Package::edit_route_slip`] stages route slip edits on an [`Edit`];
//! [`Edit::commit`] publishes a new [`crate::Package`] with an exact
//! whole-package [`Patch`] and [`Diagnostics`];
//! [`crate::Package::apply_route_slip`] replays the patch. The part-level codec
//! stays public in [`super::record`]; the byte-level editor is crate-private.

#![allow(
    unused_imports,
    reason = "the family forwards the part editor's signatures verbatim"
)]

use super::codec;
use super::model::Metadata;
use super::package::Editor;
use super::transaction::{
    Edit as RecordEdit, Error as TransactionError, Patch as RecordPatch, RecipientSelector,
    Snapshot as TransactionSnapshot,
};
use super::validation;
use crate::package::Error as PackageError;
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{EditProtection, PackagePatch, ProtectionPolicy, classify};
use litchi_ole_common::object::{Edit as ObjectEditor, Limits, Targets};

crate::scoped::family!("route slip");

/// A staged route slip edit of one package. Dropping it discards it.
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

    /// Returns the current route-slip metadata, when present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.inner.metadata()
    }

    /// Returns the current immutable route-slip semantic snapshot.
    #[must_use]
    pub fn route_slip(&self) -> &TransactionSnapshot {
        self.inner.route_slip()
    }

    /// Sets or replaces the complete route-slip metadata.
    pub fn set(&mut self, metadata: Metadata) -> Result<()> {
        self.inner.set(metadata).map(|_| ()).map_err(Error::from)
    }

    /// Replaces existing route-slip metadata, rejecting an absent range.
    pub fn replace(&mut self, metadata: Metadata) -> Result<()> {
        self.inner
            .replace(metadata)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Applies and publishes a semantic transaction atomically.
    pub fn apply(&mut self, transaction: RecordEdit) -> Result<()> {
        self.inner
            .apply(transaction)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Sets the current recipient stage and stages the edit.
    pub fn set_stage(&mut self, selector: RecipientSelector<'_>) -> Result<()> {
        self.inner
            .set_stage(selector)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Advances to the next recipient and publishes the lifecycle edit.
    pub fn advance_stage(&mut self) -> Result<()> {
        self.inner.advance_stage().map(|_| ()).map_err(Error::from)
    }

    /// Adds a recipient and stages the edit.
    pub fn add_recipient(&mut self, recipient: super::model::Recipient) -> Result<()> {
        self.inner
            .add_recipient(recipient)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Replaces a selected recipient and stages the edit.
    pub fn replace_recipient(
        &mut self,
        selector: RecipientSelector<'_>,
        recipient: super::model::Recipient,
    ) -> Result<()> {
        self.inner
            .replace_recipient(selector, recipient)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Removes a selected recipient and stages the edit.
    pub fn remove_recipient(&mut self, selector: RecipientSelector<'_>) -> Result<()> {
        self.inner
            .remove_recipient(selector)
            .map(|_| ())
            .map_err(Error::from)
    }

    /// Clears the route-slip range and its semantic metadata.
    pub fn clear(&mut self) -> Result<()> {
        self.inner.clear().map(|_| ()).map_err(Error::from)
    }

    /// Completes routing by clearing the route-slip FIB range.
    pub fn complete(&mut self) -> Result<()> {
        self.inner.complete().map(|_| ()).map_err(Error::from)
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
    /// Start a route slip edit of this package under its retained edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's route slip cannot be read for editing.
    pub fn edit_route_slip(&self) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture(self)?)
    }

    /// Start a route slip edit of this package under explicit edit options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's route slip cannot be read for editing.
    pub fn edit_route_slip_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture_with_options(self, options)?)
    }

    /// Replay a route slip patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed package.
    pub fn apply_route_slip(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }
}
