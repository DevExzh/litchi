//! Inert `ObjectPool` storage lifecycle edits.

use super::super::codec::corrupted;
use super::super::model::Editor;
use super::inventory::object_for_reference;
use crate::package::{Error as PackageError, Result};

impl Editor {
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
        let reference = self
            .objects()?
            .into_iter()
            .find(|value| value.storage_id == storage_id)
            .ok_or_else(|| corrupted("managed embedded-object field was not found"))?;
        let object = object_for_reference(self.package.objects(), &reference)
            .ok_or_else(|| corrupted("ObjectPool storage target is missing"))?;
        if object.compound() == compound_file.as_slice() {
            return Ok(());
        }

        let mut candidate = self.clone();
        candidate
            .package
            .replace(object.key(), compound_file)
            .map_err(PackageError::from)?;
        candidate.changed = true;
        *self = candidate;
        Ok(())
    }
}
