//! The embedded-object edit family (record 0351 §2.3, record 0359).
//!
//! [`crate::Package::embedded_objects`] reads the managed field and
//! `ObjectPool` inventory, [`crate::Package::edit_embedded_objects`] stages
//! lifecycle, metadata and storage edits on an [`Edit`], and
//! [`Edit::commit`] publishes a new [`crate::Package`] with an exact
//! whole-package [`Patch`] and [`Diagnostics`];
//! [`crate::Package::apply_embedded_objects`] replays the patch. Objects are
//! never activated: storages stay inert compound files.
//!
//! Selecting a discovered native storage without editor-managed field ownership
//! for a storage-ID mutation returns [`Error::Refused`] with
//! [`RefusalKind::UnsupportedContent`]. Passive reads remain available through
//! [`Snapshot::storages`].

use super::model::{Info, Inventory, Reference, WriteOptions};
use super::transaction::{Snapshot, Transaction, TransactionError};
use crate::equation::MtefEquationWriteOptions;
use litchi_core::edit::{Refusal, RefusalKind};
use litchi_ole_common::ole_streams::{self, native, presentation};

crate::scoped::family!("embedded object");

impl From<TransactionError> for Error {
    fn from(error: TransactionError) -> Self {
        match error {
            TransactionError::Invalid(error) => Self::Package(error),
            TransactionError::Conflict => Self::PatchConflict,
        }
    }
}

/// A staged embedded-object edit of one package. Dropping it discards it.
#[derive(Clone)]
pub struct Edit {
    base: crate::scoped::Base,
    inner: Option<Transaction>,
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
    /// A native storage can be valid without the editor-owned field convention.
    /// Refuse its mutation before a managed-field lookup labels it corrupted.
    fn require_managed_storage(&mut self, storage_id: u32) -> Result<()> {
        if self.inner_mut()?.has_unmanaged_storage(storage_id)? {
            return Err(Error::Refused(Refusal::new(
                RefusalKind::UnsupportedContent,
                "edit embedded object",
                "native ObjectPool storage has no editor-managed field ownership",
            )));
        }
        Ok(())
    }

    fn open(base: crate::scoped::Base) -> Result<Self> {
        Ok(Self { base, inner: None })
    }

    fn inner_mut(&mut self) -> Result<&mut Transaction> {
        if self.inner.is_none() {
            let snapshot = Snapshot::open_with_policy(
                self.base.to_vec()?,
                self.base.options().revision_limits_value(),
                self.base.options().protection_policy_value(),
            )?;
            self.inner = Some(snapshot.edit());
        }
        self.inner.as_mut().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "embedded-object edit was not initialized".into(),
            ))
        })
    }

    fn take_inner(&mut self) -> Result<Transaction> {
        self.inner_mut()?;
        self.inner.take().ok_or_else(|| {
            Error::Package(crate::Error::InvalidFormat(
                "embedded-object edit was not initialized".into(),
            ))
        })
    }

    /// Returns the current candidate inventory without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate contains an unresolved field
    /// reference or malformed recognized metadata.
    pub fn inventory(&mut self) -> Result<Inventory> {
        Ok(self.inner_mut()?.inventory()?)
    }

    /// Returns current managed field references in document order.
    ///
    /// # Errors
    ///
    /// Returns an error when a field reference has no owning `ObjectPool`
    /// storage.
    pub fn objects(&mut self) -> Result<Vec<Reference>> {
        Ok(self.inner_mut()?.objects()?)
    }

    /// Adds one field and its inert `ObjectPool` storage.
    ///
    /// # Errors
    ///
    /// Returns an error when the options, field tables, payload bounds, or
    /// candidate CFB cannot be validated atomically.
    pub fn add(&mut self, options: WriteOptions) -> Result<Reference> {
        Ok(self.inner_mut()?.add(options)?)
    }

    /// Removes one managed field and its owning `ObjectPool` storage.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage ID is not referenced or the candidate
    /// field/package rewrite fails validation.
    pub fn remove(&mut self, storage_id: u32) -> Result<Reference> {
        self.require_managed_storage(storage_id)?;
        Ok(self.inner_mut()?.remove(storage_id)?)
    }

    /// Reorders the managed embedded-object suffix by semantic storage ID.
    ///
    /// # Errors
    ///
    /// Returns an error when the IDs are not a complete valid permutation or
    /// the candidate field/package rewrite fails validation.
    pub fn reorder(&mut self, storage_ids: &[u32]) -> Result<()> {
        for &storage_id in storage_ids {
            self.require_managed_storage(storage_id)?;
        }
        Ok(self.inner_mut()?.reorder(storage_ids)?)
    }

    /// Replaces one complete standalone object CFB without activating it.
    ///
    /// The field reference and surrounding DOC streams are retained. The
    /// replacement itself is bounded and reparsed as an inert CFB before it
    /// can replace the candidate storage.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage ID is not referenced, the replacement
    /// is malformed/oversized, or the candidate package cannot be validated.
    pub fn replace_storage(&mut self, storage_id: u32, compound_file: Vec<u8>) -> Result<()> {
        self.require_managed_storage(storage_id)?;
        Ok(self
            .inner_mut()?
            .replace_storage(storage_id, compound_file)?)
    }

    /// Replaces or creates the passive `\x03ObjInfo` ODT metadata stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage ID is not referenced, the typed ODT
    /// is invalid, or the candidate package cannot be validated.
    pub fn set_info(&mut self, storage_id: u32, info: Info) -> Result<()> {
        self.require_managed_storage(storage_id)?;
        Ok(self.inner_mut()?.set_info(storage_id, info)?)
    }

    /// Clone-edits one existing, valid passive `\x03ObjInfo` ODT value.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage has no valid typed ODT or the edited
    /// candidate violates an ODT or package invariant.
    pub fn update_info<F>(&mut self, storage_id: u32, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Info),
    {
        self.require_managed_storage(storage_id)?;
        Ok(self.inner_mut()?.update_info(storage_id, edit)?)
    }

    /// Clone-edits inert OLEDS `\x01Ole` metadata without resolving its link.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage has no valid OLEDS link stream, the
    /// callback rejects the candidate, or the package cannot be validated.
    pub fn update_link<F>(
        &mut self,
        storage_id: u32,
        limits: litchi_ole_common::object::link::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(
            &mut litchi_ole_common::object::link::Link,
        ) -> std::result::Result<(), litchi_ole_common::Error>,
    {
        self.require_managed_storage(storage_id)?;
        Ok(self.inner_mut()?.update_link(storage_id, limits, edit)?)
    }

    /// Clone-edits one presentation under explicit OLEDS limits.
    pub fn update_presentation<F>(
        &mut self,
        storage_id: u32,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut presentation::Edit) -> std::result::Result<(), litchi_ole_common::Error>,
    {
        self.require_managed_storage(storage_id)?;
        Ok(self
            .inner_mut()?
            .update_presentation(storage_id, index, limits, edit)?)
    }

    /// Clone-edits a reference-selected presentation under explicit limits.
    pub fn update_presentation_for<F>(
        &mut self,
        reference: &Reference,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut presentation::Edit) -> std::result::Result<(), litchi_ole_common::Error>,
    {
        Ok(self
            .inner_mut()?
            .update_presentation_for(reference, index, limits, edit)?)
    }

    /// Applies a source-checked presentation patch to one managed object.
    pub fn apply_presentation_patch(
        &mut self,
        storage_id: u32,
        index: usize,
        patch: &presentation::Patch,
    ) -> Result<()> {
        self.require_managed_storage(storage_id)?;
        Ok(self
            .inner_mut()?
            .apply_presentation_patch(storage_id, index, patch)?)
    }

    /// Applies a presentation patch selected by a DOC object reference.
    pub fn apply_presentation_patch_for(
        &mut self,
        reference: &Reference,
        index: usize,
        patch: &presentation::Patch,
    ) -> Result<()> {
        Ok(self
            .inner_mut()?
            .apply_presentation_patch_for(reference, index, patch)?)
    }

    /// Clone-edits one native-data stream under explicit OLEDS limits.
    pub fn update_native<F>(
        &mut self,
        storage_id: u32,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut native::Edit) -> std::result::Result<(), litchi_ole_common::Error>,
    {
        self.require_managed_storage(storage_id)?;
        Ok(self.inner_mut()?.update_native(storage_id, limits, edit)?)
    }

    /// Clone-edits reference-selected native data under explicit limits.
    pub fn update_native_for<F>(
        &mut self,
        reference: &Reference,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut native::Edit) -> std::result::Result<(), litchi_ole_common::Error>,
    {
        Ok(self
            .inner_mut()?
            .update_native_for(reference, limits, edit)?)
    }

    /// Applies a source-checked native-data patch to one managed object.
    pub fn apply_native_patch(&mut self, storage_id: u32, patch: &native::Patch) -> Result<()> {
        self.require_managed_storage(storage_id)?;
        Ok(self.inner_mut()?.apply_native_patch(storage_id, patch)?)
    }

    /// Applies a native-data patch selected by a DOC object reference.
    pub fn apply_native_patch_for(
        &mut self,
        reference: &Reference,
        patch: &native::Patch,
    ) -> Result<()> {
        Ok(self.inner_mut()?.apply_native_patch_for(reference, patch)?)
    }

    /// Appends a native Equation.3 EMBED field with a real PICF bitmap
    /// preview.
    ///
    /// # Errors
    ///
    /// Returns an error when the equation, preview or candidate package
    /// cannot be validated; the staged state is unchanged.
    pub fn add_mtef_equation(&mut self, options: MtefEquationWriteOptions) -> Result<Reference> {
        Ok(self.inner_mut()?.add_mtef_equation(options)?)
    }

    /// Validate the staged edits and publish them as a new package.
    ///
    /// # Errors
    ///
    /// Returns an error when the protection policy refuses a changed write,
    /// the package cannot be republished, or [`Error::Signed`] for a changed
    /// signed package.
    pub fn commit(mut self) -> Result<Commit> {
        let after = self.take_inner()?.commit()?.into_patch().after().to_vec();
        Ok(published(self.base.publish(after)?))
    }
}

impl crate::Package {
    /// The embedded-object read view of this package: its managed field and
    /// `ObjectPool` inventory, presentations, native data and transfer
    /// exports.
    ///
    /// # Errors
    ///
    /// Returns an error when the field or `ObjectPool` tables cannot be read.
    pub fn embedded_objects(&self) -> Result<Snapshot> {
        Ok(Snapshot::open_with_policy(
            self.to_bytes()?,
            self.retained_edit_options().revision_limits_value(),
            self.retained_edit_options().protection_policy_value(),
        )?)
    }

    /// Start an embedded-object edit of this package under its retained edit
    /// options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's field or `ObjectPool` tables
    /// cannot be read for editing.
    pub fn edit_embedded_objects(&self) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture(self)?)
    }

    /// Start an embedded-object edit of this package under explicit edit
    /// options.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's field or `ObjectPool` tables
    /// cannot be read for editing.
    pub fn edit_embedded_objects_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Edit::open(crate::scoped::Base::capture_with_options(self, options)?)
    }

    /// Replay an embedded-object patch onto the exact package it was made
    /// from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed package.
    pub fn apply_embedded_objects(&self, patch: &Patch) -> Result<Commit> {
        apply_patch(self, patch)
    }
}
