//! Cross-record and collection-level invariants.

use super::super::codec::corrupted;
use super::super::model::{Collection, ExternalObject, MAX_OLE_OBJECTS, UnknownRecord};
use crate::package::{Error, Result};
use crate::persist::PersistMapping;
use std::collections::HashSet;

impl UnknownRecord {
    pub(crate) fn validate_for(&self, object_count: usize) -> Result<()> {
        if self.object_index > object_count {
            return corrupted("unknown ExObjList record has an invalid source slot");
        }
        let expected_length = usize::try_from(self.record.data_length)
            .map_err(|_err| Error::Corrupted("unknown ExObjList record size overflows".into()))?;
        if expected_length != self.record.data.len() {
            return corrupted("unknown ExObjList record has inconsistent payload length");
        }
        self.to_record_bytes().map(|_| ())
    }
}

impl Collection {
    /// Validate the collection's seed, ID, and unknown-record invariants.
    ///
    /// # Errors
    ///
    /// Returns an error if the identifier seed or any object ID is out of
    /// range, if object IDs are duplicated, if the object count exceeds the
    /// limit, or if a preserved unknown record is inconsistent.
    pub fn validate(&self) -> Result<()> {
        validate_collection(self.id_seed, &self.objects)?;
        for record in &self.unknown_records {
            record.validate_for(self.objects.len())?;
        }
        Ok(())
    }

    /// Validate that every object's persist ID is present in `mapping`.
    ///
    /// # Errors
    ///
    /// Returns an error if any object references a persist ID that has no
    /// offset in the mapping.
    pub fn validate_persist_mapping(&self, mapping: &PersistMapping) -> Result<()> {
        for object in &self.objects {
            let id = object.persist_id();
            if mapping.get_offset(id).is_none() {
                return corrupted(format!("OLE object references missing persist ID {id}"));
            }
        }
        Ok(())
    }
}

pub(crate) fn validate_collection(id_seed: u32, objects: &[ExternalObject]) -> Result<()> {
    if id_seed == 0 || id_seed > i32::MAX as u32 {
        return corrupted("ExObjList identifier seed must fit a positive signed integer");
    }
    if objects.len() > MAX_OLE_OBJECTS {
        return corrupted(format!(
            "external-object list exceeds {MAX_OLE_OBJECTS} OLE objects"
        ));
    }
    let mut ids = HashSet::new();
    for object in objects {
        let id = object.id();
        if id == 0 || id > id_seed {
            return corrupted(format!(
                "OLE object ID {id} is zero or exceeds ExObjList seed {id_seed}"
            ));
        }
        if object.persist_id() == 0 {
            return corrupted(format!("OLE object ID {id} has zero persist ID"));
        }
        if !ids.insert(id) {
            return corrupted(format!(
                "external-object list contains duplicate OLE object ID {id}"
            ));
        }
    }
    Ok(())
}
