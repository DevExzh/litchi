//! Failure-atomic transactions over DOC embedded-object snapshots.

use super::super::model::{Editor, Info, Inventory, Reference, WriteOptions};
use super::{Commit, Patch, Snapshot};
use crate::package::Error as PackageError;
use litchi_ole_common::ole_streams::{self, native, presentation};
use std::fmt;

/// A staged, clone-first DOC embedded-object edit.
///
/// The source snapshot remains immutable. Every semantic operation is
/// validated against a private candidate before it replaces the transaction's
/// editor, so a failed lifecycle or metadata operation cannot partially
/// publish a field, `ObjectPool` storage, or opaque stream change.
#[derive(Clone)]
pub(crate) struct Transaction {
    source: Snapshot,
    editor: Editor,
}

impl Transaction {
    pub(super) fn new(source: Snapshot) -> Self {
        Self {
            editor: source.editor().clone(),
            source,
        }
    }

    /// Returns the immutable source snapshot used for stale-source checks.
    #[must_use]
    pub(crate) const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Materializes the current candidate as a validated snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate cannot be rendered or fails DOC,
    /// CFB, resource-bound, or owner/reference validation.
    pub(crate) fn snapshot(&self) -> Result<Snapshot, TransactionError> {
        let bytes = self
            .editor
            .clone()
            .finish()
            .map_err(TransactionError::Invalid)?;
        if bytes == self.source.source_bytes() {
            return Ok(self.source.clone());
        }
        Snapshot::open_with_policy(bytes, self.source.limits(), self.source.protection_policy())
            .map_err(TransactionError::Invalid)
    }

    /// Returns the current candidate inventory without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate contains an unresolved field
    /// reference or malformed recognized metadata.
    pub(crate) fn inventory(&self) -> Result<Inventory, TransactionError> {
        self.editor.inventory().map_err(TransactionError::Invalid)
    }

    /// Returns current managed field references in document order.
    ///
    /// # Errors
    ///
    /// Returns an error when a field reference has no owning `ObjectPool`
    /// storage.
    pub(crate) fn objects(&self) -> Result<Vec<Reference>, TransactionError> {
        self.editor.objects().map_err(TransactionError::Invalid)
    }

    /// Adds one field and its inert `ObjectPool` storage.
    ///
    /// # Errors
    ///
    /// Returns an error when the options, field tables, payload bounds, or
    /// candidate CFB cannot be validated atomically.
    pub(crate) fn add(&mut self, options: WriteOptions) -> Result<Reference, TransactionError> {
        self.editor.add(options).map_err(TransactionError::Invalid)
    }

    /// Appends a native Equation.3 EMBED field with a real PICF preview.
    pub(crate) fn add_mtef_equation(
        &mut self,
        options: crate::equation::MtefEquationWriteOptions,
    ) -> Result<Reference, TransactionError> {
        self.editor
            .add_mtef_equation(options)
            .map_err(TransactionError::Invalid)
    }

    /// Removes one managed field and its owning `ObjectPool` storage.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage ID is not referenced or the candidate
    /// field/package rewrite fails validation.
    pub(crate) fn remove(&mut self, storage_id: u32) -> Result<Reference, TransactionError> {
        self.editor
            .remove(storage_id)
            .map_err(TransactionError::Invalid)
    }

    /// Reorders the managed embedded-object suffix by semantic storage ID.
    ///
    /// # Errors
    ///
    /// Returns an error when the IDs are not a complete valid permutation or
    /// the candidate field/package rewrite fails validation.
    pub(crate) fn reorder(&mut self, storage_ids: &[u32]) -> Result<(), TransactionError> {
        self.editor
            .reorder(storage_ids)
            .map_err(TransactionError::Invalid)
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
    pub(crate) fn replace_storage(
        &mut self,
        storage_id: u32,
        compound_file: Vec<u8>,
    ) -> Result<(), TransactionError> {
        self.editor
            .replace_storage(storage_id, compound_file)
            .map_err(TransactionError::Invalid)
    }

    /// Replaces or creates the passive `\x03ObjInfo` ODT metadata stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage ID is not referenced, the typed ODT
    /// is invalid, or the candidate package cannot be validated.
    pub(crate) fn set_info(&mut self, storage_id: u32, info: Info) -> Result<(), TransactionError> {
        self.editor
            .set_info(storage_id, info)
            .map_err(TransactionError::Invalid)
    }

    /// Clone-edits one existing, valid passive `\x03ObjInfo` ODT value.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage has no valid typed ODT or the edited
    /// candidate violates an ODT or package invariant.
    pub(crate) fn update_info<F>(
        &mut self,
        storage_id: u32,
        edit: F,
    ) -> Result<(), TransactionError>
    where
        F: FnOnce(&mut Info),
    {
        self.editor
            .update_info(storage_id, edit)
            .map_err(TransactionError::Invalid)
    }

    /// Clone-edits inert OLEDS `\x01Ole` metadata without resolving its link.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage has no valid OLEDS link stream, the
    /// callback rejects the candidate, or the package cannot be validated.
    pub(crate) fn update_link<F>(
        &mut self,
        storage_id: u32,
        limits: litchi_ole_common::object::link::Limits,
        edit: F,
    ) -> Result<(), TransactionError>
    where
        F: FnOnce(
            &mut litchi_ole_common::object::link::Link,
        ) -> Result<(), litchi_ole_common::Error>,
    {
        self.editor
            .update_link(storage_id, limits, edit)
            .map_err(TransactionError::Invalid)
    }

    /// Clone-edits one presentation under explicit OLEDS limits.
    pub(crate) fn update_presentation<F>(
        &mut self,
        storage_id: u32,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), TransactionError>
    where
        F: FnOnce(&mut presentation::Edit) -> Result<(), litchi_ole_common::Error>,
    {
        self.editor
            .update_presentation(storage_id, index, limits, edit)
            .map_err(TransactionError::Invalid)
    }

    /// Clone-edits a reference-selected presentation under explicit limits.
    pub(crate) fn update_presentation_for<F>(
        &mut self,
        reference: &Reference,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), TransactionError>
    where
        F: FnOnce(&mut presentation::Edit) -> Result<(), litchi_ole_common::Error>,
    {
        self.editor
            .update_presentation_for(reference, index, limits, edit)
            .map_err(TransactionError::Invalid)
    }

    /// Applies a source-checked presentation patch to one managed object.
    pub(crate) fn apply_presentation_patch(
        &mut self,
        storage_id: u32,
        index: usize,
        patch: &presentation::Patch,
    ) -> Result<(), TransactionError> {
        self.editor
            .apply_presentation_patch(storage_id, index, patch)
            .map_err(TransactionError::Invalid)
    }

    /// Applies a presentation patch selected by a DOC object reference.
    pub(crate) fn apply_presentation_patch_for(
        &mut self,
        reference: &Reference,
        index: usize,
        patch: &presentation::Patch,
    ) -> Result<(), TransactionError> {
        self.editor
            .apply_presentation_patch_for(reference, index, patch)
            .map_err(TransactionError::Invalid)
    }

    /// Clone-edits one native-data stream under explicit OLEDS limits.
    pub(crate) fn update_native<F>(
        &mut self,
        storage_id: u32,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), TransactionError>
    where
        F: FnOnce(&mut native::Edit) -> Result<(), litchi_ole_common::Error>,
    {
        self.editor
            .update_native(storage_id, limits, edit)
            .map_err(TransactionError::Invalid)
    }

    /// Clone-edits reference-selected native data under explicit limits.
    pub(crate) fn update_native_for<F>(
        &mut self,
        reference: &Reference,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), TransactionError>
    where
        F: FnOnce(&mut native::Edit) -> Result<(), litchi_ole_common::Error>,
    {
        self.editor
            .update_native_for(reference, limits, edit)
            .map_err(TransactionError::Invalid)
    }

    /// Applies a source-checked native-data patch to one managed object.
    pub(crate) fn apply_native_patch(
        &mut self,
        storage_id: u32,
        patch: &native::Patch,
    ) -> Result<(), TransactionError> {
        self.editor
            .apply_native_patch(storage_id, patch)
            .map_err(TransactionError::Invalid)
    }

    /// Applies a native-data patch selected by a DOC object reference.
    pub(crate) fn apply_native_patch_for(
        &mut self,
        reference: &Reference,
        patch: &native::Patch,
    ) -> Result<(), TransactionError> {
        self.editor
            .apply_native_patch_for(reference, patch)
            .map_err(TransactionError::Invalid)
    }

    /// Whether the current candidate serializes differently from its source.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate cannot be rendered for comparison.
    pub(crate) fn is_changed(&self) -> Result<bool, TransactionError> {
        self.editor
            .clone()
            .finish()
            .map(|bytes| bytes != self.source.source_bytes())
            .map_err(TransactionError::Invalid)
    }

    /// Publishes the candidate as a source-checked reversible commit.
    ///
    /// # Errors
    ///
    /// Returns an error when the candidate cannot be rendered or the resulting
    /// snapshot fails bounded DOC/CFB validation.
    pub(crate) fn commit(self) -> Result<Commit, TransactionError> {
        let before = self.source;
        let after = self.editor.finish().map_err(TransactionError::Invalid)?;
        let patch = Patch::new(before.source_bytes().to_vec(), after.clone());
        let snapshot = if patch.is_noop() {
            before
        } else {
            Snapshot::open_with_policy(after, before.limits(), before.protection_policy())
                .map_err(TransactionError::Invalid)?
        };
        Ok(Commit::new(snapshot, patch))
    }
}

/// Errors produced by a DOC embedded-object transaction.
#[derive(Debug)]
pub enum TransactionError {
    /// The candidate or serialized DOC violates a bounded format invariant.
    Invalid(PackageError),
    /// The patch was applied to a snapshot other than its exact source.
    Conflict,
}

impl fmt::Display for TransactionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Conflict => formatter.write_str("embedded-object snapshot conflict"),
        }
    }
}

impl std::error::Error for TransactionError {}

impl From<PackageError> for TransactionError {
    fn from(error: PackageError) -> Self {
        Self::Invalid(error)
    }
}

impl From<TransactionError> for PackageError {
    fn from(source_error: TransactionError) -> Self {
        match source_error {
            TransactionError::Invalid(error) => error,
            TransactionError::Conflict => {
                PackageError::InvalidFormat("embedded-object snapshot conflict".into())
            },
        }
    }
}
