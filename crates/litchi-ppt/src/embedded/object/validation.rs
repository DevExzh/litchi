//! Transactional validation and mutation for OLE object snapshots.

use super::codec::corrupted;
use super::model::*;
use crate::package::{Error, Result};
use crate::persist::PersistMapping;
use std::collections::HashSet;

impl UnknownRecord {
    pub(crate) fn validate_for(&self, object_count: usize) -> Result<()> {
        if self.object_index > object_count {
            return corrupted("unknown ExObjList record has an invalid source slot");
        }
        let expected_length = usize::try_from(self.record.data_length)
            .map_err(|_| Error::Corrupted("unknown ExObjList record size overflows".into()))?;
        if expected_length != self.record.data.len() {
            return corrupted("unknown ExObjList record has inconsistent payload length");
        }
        self.to_record_bytes().map(|_| ())
    }
}

impl Collection {
    pub fn add(&mut self, object: ExternalObject) -> Result<()> {
        let mut candidate = self.clone();
        let insertion_index = candidate.objects.len();
        candidate.objects.push(object);
        for record in &mut candidate.unknown_records {
            if record.object_index >= insertion_index {
                record.object_index += 1;
            }
        }
        validate_collection(candidate.id_seed, &candidate.objects)?;
        *self = candidate;
        Ok(())
    }

    pub fn update<F>(&mut self, id: u32, edit: F) -> Result<()>
    where
        F: FnOnce(&mut ExternalObject) -> Result<()>,
    {
        let mut candidate = self.clone();
        let object = candidate
            .objects
            .iter_mut()
            .find(|object| object.id() == id)
            .ok_or_else(|| Error::Corrupted(format!("OLE object ID {id} was not found")))?;
        edit(object)?;
        validate_collection(candidate.id_seed, &candidate.objects)?;
        *self = candidate;
        Ok(())
    }

    pub fn replace(&mut self, id: u32, replacement: ExternalObject) -> Result<ExternalObject> {
        let mut candidate = self.clone();
        let index = candidate
            .objects
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| Error::Corrupted(format!("OLE object ID {id} was not found")))?;
        let previous = std::mem::replace(&mut candidate.objects[index], replacement);
        validate_collection(candidate.id_seed, &candidate.objects)?;
        *self = candidate;
        Ok(previous)
    }

    pub fn remove(&mut self, id: u32) -> Result<ExternalObject> {
        let mut candidate = self.clone();
        let index = candidate
            .objects
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| Error::Corrupted(format!("OLE object ID {id} was not found")))?;
        let removed = candidate.objects.remove(index);
        for record in &mut candidate.unknown_records {
            if record.object_index > index {
                record.object_index -= 1;
            }
        }
        *self = candidate;
        Ok(removed)
    }

    pub fn reorder(&mut self, ids: &[u32]) -> Result<()> {
        if ids.len() != self.objects.len() {
            return corrupted("OLE reorder must contain every object exactly once");
        }
        let mut remaining = self.objects.clone();
        let mut candidate = Vec::with_capacity(ids.len());
        for id in ids {
            let index = remaining
                .iter()
                .position(|object| object.id() == *id)
                .ok_or_else(|| {
                    Error::Corrupted(format!("unknown or repeated OLE object ID {id}"))
                })?;
            candidate.push(remaining.remove(index));
        }
        validate_collection(self.id_seed, &candidate)?;
        self.objects = candidate;
        Ok(())
    }

    pub fn validate(&self) -> Result<()> {
        validate_collection(self.id_seed, &self.objects)?;
        for record in &self.unknown_records {
            record.validate_for(self.objects.len())?;
        }
        Ok(())
    }

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
