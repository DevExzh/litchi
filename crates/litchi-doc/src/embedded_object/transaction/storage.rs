//! Inert `ObjectPool` storage lifecycle edits.

use super::super::codec::corrupted;
use super::super::model::{Editor, Reference};
use super::super::storage::OBJECT_POOL;
use crate::package::{Error as PackageError, Result};
use litchi_cfb::OleError;
use litchi_ole_common::object::Object;
use litchi_ole_common::ole_streams::{
    self, NativePatch, NativeSnapshot, NativeTransaction, PresentationPatch, PresentationSnapshot,
    PresentationTransaction,
};

fn reference_for_storage_id(editor: &Editor, storage_id: u32) -> Result<Reference> {
    editor
        .objects()?
        .into_iter()
        .find(|reference| reference.storage_id == storage_id)
        .ok_or_else(|| corrupted("managed embedded-object field was not found"))
}

fn object_for_reference_checked<'a>(
    editor: &'a Editor,
    reference: &Reference,
) -> Result<&'a Object> {
    let object = editor
        .package
        .objects()
        .get(&reference.storage_name)
        .ok_or_else(|| {
            corrupted(format!(
                "ObjectPool storage target {:?} is missing",
                reference.storage_name
            ))
        })?;

    let path = object.path();
    if path.len() != 2
        || path[0] != OBJECT_POOL
        || path[1] != reference.storage_name
        || object.key() != reference.storage_name
    {
        return Err(corrupted(
            "embedded-object reference does not identify its ObjectPool target",
        ));
    }

    let numeric_id = reference
        .storage_name
        .strip_prefix('_')
        .filter(|digits| !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()))
        .and_then(|digits| digits.parse::<u32>().ok())
        .ok_or_else(|| corrupted("embedded-object storage name has no valid numeric identity"))?;
    if numeric_id != reference.storage_id {
        return Err(corrupted(
            "embedded-object reference storage name and ID disagree",
        ));
    }

    if !editor.objects()?.into_iter().any(|managed| {
        managed.storage_id == reference.storage_id && managed.storage_name == reference.storage_name
    }) {
        return Err(corrupted(
            "embedded-object reference is not owned by a managed DOC field",
        ));
    }

    let numeric_target_count = editor
        .package
        .objects()
        .iter()
        .filter(|candidate| {
            let path = candidate.path();
            path.len() == 2
                && path[0] == OBJECT_POOL
                && path[1]
                    .strip_prefix('_')
                    .filter(|digits| {
                        !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit())
                    })
                    .and_then(|digits| digits.parse::<u32>().ok())
                    == Some(numeric_id)
        })
        .count();
    if numeric_target_count != 1 {
        return Err(corrupted(format!(
            "ObjectPool storage ID {} is ambiguous",
            numeric_id
        )));
    }

    Ok(object)
}

fn missing_stream() -> PackageError {
    PackageError::from(OleError::StreamNotFound)
}

impl Editor {
    /// Returns one managed object's inert OLEDS presentation stream.
    ///
    /// The DOC field and `ObjectPool` storage are resolved by the DOC owner;
    /// the selected stream is then parsed by the shared bounded OLEDS owner.
    /// No presentation data is rendered or activated.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage is not referenced, the stream index
    /// or payload is malformed, or the selected OLEDS limits are exceeded.
    pub fn presentation(
        &self,
        storage_id: u32,
        index: usize,
    ) -> Result<Option<PresentationSnapshot>> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.presentation_for(&reference, index)
    }

    /// Returns one managed object's OLEDS presentation under explicit limits.
    pub fn presentation_with_limits(
        &self,
        storage_id: u32,
        index: usize,
        limits: ole_streams::Limits,
    ) -> Result<Option<PresentationSnapshot>> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.presentation_for_with_limits(&reference, index, limits)
    }

    /// Returns one managed object's presentation selected by its DOC reference.
    pub fn presentation_for(
        &self,
        reference: &Reference,
        index: usize,
    ) -> Result<Option<PresentationSnapshot>> {
        self.presentation_for_with_limits(reference, index, ole_streams::Limits::default())
    }

    /// Returns one DOC-reference-selected presentation under explicit limits.
    pub fn presentation_for_with_limits(
        &self,
        reference: &Reference,
        index: usize,
        limits: ole_streams::Limits,
    ) -> Result<Option<PresentationSnapshot>> {
        object_for_reference_checked(self, reference)?
            .presentation_with_limits(index, limits)
            .map_err(PackageError::from)
    }

    /// Returns all managed OLEDS presentations in numeric order.
    pub fn presentations(&self, storage_id: u32) -> Result<Vec<(usize, PresentationSnapshot)>> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.presentations_for(&reference)
    }

    /// Returns all managed presentations under explicit limits.
    pub fn presentations_with_limits(
        &self,
        storage_id: u32,
        limits: ole_streams::Limits,
    ) -> Result<Vec<(usize, PresentationSnapshot)>> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.presentations_for_with_limits(&reference, limits)
    }

    /// Returns all presentations selected by a DOC object reference.
    pub fn presentations_for(
        &self,
        reference: &Reference,
    ) -> Result<Vec<(usize, PresentationSnapshot)>> {
        self.presentations_for_with_limits(reference, ole_streams::Limits::default())
    }

    /// Returns all DOC-reference-selected presentations under explicit limits.
    pub fn presentations_for_with_limits(
        &self,
        reference: &Reference,
        limits: ole_streams::Limits,
    ) -> Result<Vec<(usize, PresentationSnapshot)>> {
        object_for_reference_checked(self, reference)?
            .presentations_with_limits(limits)
            .map_err(PackageError::from)
    }

    /// Returns one managed object's inert OLEDS native-data stream.
    pub fn native(&self, storage_id: u32) -> Result<Option<NativeSnapshot>> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.native_for(&reference)
    }

    /// Returns one managed object's native-data stream under explicit limits.
    pub fn native_with_limits(
        &self,
        storage_id: u32,
        limits: ole_streams::Limits,
    ) -> Result<Option<NativeSnapshot>> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.native_for_with_limits(&reference, limits)
    }

    /// Returns native data selected by a DOC object reference.
    pub fn native_for(&self, reference: &Reference) -> Result<Option<NativeSnapshot>> {
        self.native_for_with_limits(reference, ole_streams::Limits::default())
    }

    /// Returns DOC-reference-selected native data under explicit limits.
    pub fn native_for_with_limits(
        &self,
        reference: &Reference,
        limits: ole_streams::Limits,
    ) -> Result<Option<NativeSnapshot>> {
        object_for_reference_checked(self, reference)?
            .native_with_limits(limits)
            .map_err(PackageError::from)
    }

    /// Replaces one complete managed `ObjectPool` storage with a bounded,
    /// standalone CFB payload.
    ///
    /// The DOC field reference, WordDocument/Table/Data streams, and every
    /// other package entry remain untouched. The replacement CFB is only
    /// parsed and copied as opaque storage; no class, moniker, macro, or
    /// payload execution path is entered.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage ID is not referenced, the replacement
    /// is malformed/oversized, or the candidate package cannot be validated.
    pub fn replace_storage(&mut self, storage_id: u32, compound_file: Vec<u8>) -> Result<()> {
        let reference = reference_for_storage_id(self, storage_id)?;
        let object = object_for_reference_checked(self, &reference)?;
        let key = object.key().to_owned();
        let Some(prepared) = self
            .package
            .prepare_replacement(&key, compound_file)
            .map_err(PackageError::from)?
        else {
            return Ok(());
        };

        let mut candidate = self.clone();
        candidate
            .package
            .replace_prepared(prepared)
            .map_err(PackageError::from)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Atomically edits one managed OLEDS presentation stream.
    pub fn update_presentation<F>(&mut self, storage_id: u32, index: usize, edit: F) -> Result<()>
    where
        F: FnOnce(&mut PresentationTransaction) -> std::result::Result<(), OleError>,
    {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.update_presentation_for(&reference, index, edit)
    }

    /// Atomically edits a DOC-reference-selected OLEDS presentation stream.
    pub fn update_presentation_for<F>(
        &mut self,
        reference: &Reference,
        index: usize,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut PresentationTransaction) -> std::result::Result<(), OleError>,
    {
        self.update_presentation_for_with_limits(
            reference,
            index,
            ole_streams::Limits::default(),
            edit,
        )
    }

    /// Atomically edits a presentation under explicit OLEDS limits.
    pub fn update_presentation_with_limits<F>(
        &mut self,
        storage_id: u32,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut PresentationTransaction) -> std::result::Result<(), OleError>,
    {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.update_presentation_for_with_limits(&reference, index, limits, edit)
    }

    /// Atomically edits a DOC-reference-selected presentation under limits.
    pub fn update_presentation_for_with_limits<F>(
        &mut self,
        reference: &Reference,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut PresentationTransaction) -> std::result::Result<(), OleError>,
    {
        let object = object_for_reference_checked(self, reference)?;
        let source = object
            .presentation_with_limits(index, limits)
            .map_err(PackageError::from)?
            .ok_or_else(missing_stream)?;
        let key = object.key().to_owned();
        let mut candidate = self.clone();
        candidate
            .package
            .update_presentation_with_limits(&key, index, limits, edit)
            .map_err(PackageError::from)?;
        let replacement = candidate
            .presentation_for_with_limits(reference, index, limits)?
            .ok_or_else(missing_stream)?;
        if source.bytes() == replacement.bytes() {
            return Ok(());
        }
        candidate.validate_references()?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Applies a source-checked OLEDS presentation patch atomically.
    pub fn apply_presentation_patch(
        &mut self,
        storage_id: u32,
        index: usize,
        patch: &PresentationPatch,
    ) -> Result<()> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.apply_presentation_patch_for(&reference, index, patch)
    }

    /// Applies a presentation patch selected by a DOC object reference.
    pub fn apply_presentation_patch_for(
        &mut self,
        reference: &Reference,
        index: usize,
        patch: &PresentationPatch,
    ) -> Result<()> {
        let object = object_for_reference_checked(self, reference)?;
        let source = object
            .presentation_with_limits(index, patch.source().limits())
            .map_err(PackageError::from)?
            .ok_or_else(missing_stream)?;
        patch.apply(&source).map_err(PackageError::from)?;
        if patch.is_noop() {
            return Ok(());
        }
        let key = object.key().to_owned();
        let mut candidate = self.clone();
        candidate
            .package
            .apply_presentation_patch(&key, index, patch)
            .map_err(PackageError::from)?;
        candidate.validate_references()?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Atomically edits one managed OLEDS native-data stream.
    pub fn update_native<F>(&mut self, storage_id: u32, edit: F) -> Result<()>
    where
        F: FnOnce(&mut NativeTransaction) -> std::result::Result<(), OleError>,
    {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.update_native_for(&reference, edit)
    }

    /// Atomically edits native data selected by a DOC object reference.
    pub fn update_native_for<F>(&mut self, reference: &Reference, edit: F) -> Result<()>
    where
        F: FnOnce(&mut NativeTransaction) -> std::result::Result<(), OleError>,
    {
        self.update_native_for_with_limits(reference, ole_streams::Limits::default(), edit)
    }

    /// Atomically edits native data under explicit OLEDS limits.
    pub fn update_native_with_limits<F>(
        &mut self,
        storage_id: u32,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut NativeTransaction) -> std::result::Result<(), OleError>,
    {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.update_native_for_with_limits(&reference, limits, edit)
    }

    /// Atomically edits DOC-reference-selected native data under limits.
    pub fn update_native_for_with_limits<F>(
        &mut self,
        reference: &Reference,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<()>
    where
        F: FnOnce(&mut NativeTransaction) -> std::result::Result<(), OleError>,
    {
        let object = object_for_reference_checked(self, reference)?;
        let source = object
            .native_with_limits(limits)
            .map_err(PackageError::from)?
            .ok_or_else(missing_stream)?;
        let key = object.key().to_owned();
        let mut candidate = self.clone();
        candidate
            .package
            .update_native_with_limits(&key, limits, edit)
            .map_err(PackageError::from)?;
        let replacement = candidate
            .native_for_with_limits(reference, limits)?
            .ok_or_else(missing_stream)?;
        if source.bytes() == replacement.bytes() {
            return Ok(());
        }
        candidate.validate_references()?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Applies a source-checked OLEDS native-data patch atomically.
    pub fn apply_native_patch(&mut self, storage_id: u32, patch: &NativePatch) -> Result<()> {
        let reference = reference_for_storage_id(self, storage_id)?;
        self.apply_native_patch_for(&reference, patch)
    }

    /// Applies a native-data patch selected by a DOC object reference.
    pub fn apply_native_patch_for(
        &mut self,
        reference: &Reference,
        patch: &NativePatch,
    ) -> Result<()> {
        let object = object_for_reference_checked(self, reference)?;
        let source = object
            .native_with_limits(patch.source().limits())
            .map_err(PackageError::from)?
            .ok_or_else(missing_stream)?;
        patch.apply(&source).map_err(PackageError::from)?;
        if patch.is_noop() {
            return Ok(());
        }
        let key = object.key().to_owned();
        let mut candidate = self.clone();
        candidate
            .package
            .apply_native_patch(&key, patch)
            .map_err(PackageError::from)?;
        candidate.validate_references()?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }
}
