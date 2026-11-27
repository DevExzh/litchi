//! Transactional edits to the in-memory OLE object-list snapshot.

use super::super::codec::corrupted;
use super::super::model::{Collection, ExternalObject};
use super::collection::validate_collection;
use crate::package::{Error, Result};

impl Collection {
    /// Append `object`, keeping preserved unknown records at their positions.
    ///
    /// The collection is left unchanged when validation fails.
    ///
    /// # Errors
    ///
    /// Returns an error if the extended collection violates the seed, ID
    /// uniqueness, or object-count invariants.
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

    /// Apply `edit` to the object with the given `id`.
    ///
    /// The collection is left unchanged when the edit or validation fails.
    ///
    /// # Errors
    ///
    /// Returns an error if no object has the given `id`, if `edit` fails, or
    /// if the edited collection violates its invariants.
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

    /// Replace the object with the given `id`, returning the previous one.
    ///
    /// The collection is left unchanged when validation fails.
    ///
    /// # Errors
    ///
    /// Returns an error if no object has the given `id` or if the replaced
    /// collection violates its invariants.
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

    /// Remove the object with the given `id`, returning it.
    ///
    /// # Errors
    ///
    /// Returns an error if no object has the given `id`.
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

    /// Reorder the objects to match the sequence of `ids`.
    ///
    /// # Errors
    ///
    /// Returns an error if `ids` does not contain every object exactly once
    /// or if the reordered collection violates its invariants.
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
}
