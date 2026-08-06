//! Passive DOC ODT metadata edits inside `ObjectPool` storages.

use super::super::codec::{OBJ_INFO_STREAM, corrupted};
use super::super::model::{Editor, Info};
use super::inventory::object_for_reference;
use crate::package::{Error as PackageError, Result};
use litchi_ole_common::object::link::{Link, NAME as OLE_STREAM};
use std::result::Result as StdResult;

impl Editor {
    /// Replaces or creates the typed `\x03ObjInfo` ODT stream for one managed
    /// object.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage ID is not referenced, existing
    /// malformed `ObjInfo` bytes cannot be retained safely, the typed ODT is
    /// invalid, or the candidate CFB cannot be validated.
    ///
    /// Existing malformed `ObjInfo` bytes are rejected rather than silently
    /// overwritten; callers can preserve and inspect them through the
    /// inventory's [`crate::embedded_object::Unknown`] view. Every other stream,
    /// including the embedded binary payload, stays source-identical.
    pub fn set_info(&mut self, storage_id: u32, info: Info) -> Result<()> {
        let reference = self
            .objects()?
            .into_iter()
            .find(|value| value.storage_id == storage_id)
            .ok_or_else(|| corrupted("managed embedded-object field was not found"))?;
        let object = object_for_reference(self.package.objects(), &reference)
            .ok_or_else(|| corrupted("ObjectPool storage target is missing"))?;
        let mut path = object.path().to_vec();
        path.push(OBJ_INFO_STREAM.to_owned());
        let bytes = info.to_bytes()?;

        let mut candidate = self.clone();
        if let Some(current) = candidate.package.stream(&path).map(<[u8]>::to_vec) {
            // A typed edit must not erase malformed known bytes that are
            // currently exposed as an opaque Unknown stream.
            Info::read(&current)?;
            if current.as_slice() == bytes.as_slice() {
                return Ok(());
            }
            candidate
                .package
                .put_stream(&path, bytes)
                .map_err(PackageError::from)?;
        } else {
            candidate
                .package
                .add_stream(path, bytes)
                .map_err(PackageError::from)?;
        }
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }

    /// Clone-edits one existing valid typed `ObjInfo` stream.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage has no valid typed ODT or the edited
    /// candidate violates an ODT or package invariant.
    pub fn update_info<F>(&mut self, storage_id: u32, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Info),
    {
        let reference = self
            .objects()?
            .into_iter()
            .find(|value| value.storage_id == storage_id)
            .ok_or_else(|| corrupted("managed embedded-object field was not found"))?;
        let object = object_for_reference(self.package.objects(), &reference)
            .ok_or_else(|| corrupted("ObjectPool storage target is missing"))?;
        let mut path = object.path().to_vec();
        path.push(OBJ_INFO_STREAM.to_owned());
        let current = self
            .package
            .stream(&path)
            .map(<[u8]>::to_vec)
            .ok_or_else(|| corrupted("ObjInfo stream is missing"))?;
        let mut info = Info::read(&current)?;
        edit(&mut info);
        self.set_info(storage_id, info)
    }

    /// Clone-edits the inert OLEDS `\x01Ole` metadata stream.
    ///
    /// The callback can change flags, cache hints, class IDs, or timestamps
    /// through the shared typed link model. Moniker bytes, unknown flags,
    /// reserved fields, trailing data, and every non-link stream remain
    /// untouched. The callback never resolves or opens the link target.
    ///
    /// # Errors
    ///
    /// Returns an error when the storage has no valid OLEDS link stream, the
    /// callback rejects the candidate, or the package cannot be validated.
    pub fn update_link<F>(&mut self, storage_id: u32, edit: F) -> Result<()>
    where
        F: FnOnce(&mut Link) -> StdResult<(), litchi_cfb::OleError>,
    {
        let reference = self
            .objects()?
            .into_iter()
            .find(|value| value.storage_id == storage_id)
            .ok_or_else(|| corrupted("managed embedded-object field was not found"))?;
        let object = object_for_reference(self.package.objects(), &reference)
            .ok_or_else(|| corrupted("ObjectPool storage target is missing"))?;
        let mut path = object.path().to_vec();
        path.push(OLE_STREAM.to_owned());
        let before = self
            .package
            .stream(&path)
            .ok_or_else(|| corrupted("Ole stream is missing"))?
            .to_vec();

        let mut candidate = self.clone();
        candidate
            .package
            .update_link(object.key(), edit)
            .map_err(PackageError::from)?;
        let after = candidate
            .package
            .stream(&path)
            .ok_or_else(|| corrupted("Ole stream disappeared during edit"))?;
        if before == after {
            return Ok(());
        }
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }
}
