//! External-object collection mutations and storage staging.

use super::super::{Editor, ExternalObject, Result};
use crate::embedded::storage::Storage;
use crate::package::Error;

#[allow(
    clippy::needless_pass_by_value,
    reason = "the public `Editor` API hands ownership of the storage to the staging area; it is serialized immediately and not retained"
)]
pub(crate) fn add(
    editor: &mut Editor,
    mut object: ExternalObject,
    storage: Storage,
) -> Result<u32> {
    let persist_id = next_persist_id(editor)?;
    set_persist_id(&mut object, persist_id);

    let mut candidate = editor.clone();
    candidate.collection.add(object)?;
    candidate
        .staged_storage
        .insert(persist_id, storage.to_record_bytes()?);
    candidate.rewrite_object_list = true;
    candidate.changed = true;
    *editor = candidate;
    Ok(persist_id)
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "the public `Editor` API hands ownership of the storage to the staging area; it is serialized immediately and not retained"
)]
pub(crate) fn replace_storage(
    editor: &mut Editor,
    persist_id: u32,
    storage: Storage,
) -> Result<()> {
    if !editor
        .collection
        .objects
        .iter()
        .any(|object| object.persist_id() == persist_id)
    {
        return Err(Error::Corrupted(
            "persist ID has no OLE object reference".into(),
        ));
    }
    editor
        .staged_storage
        .insert(persist_id, storage.to_record_bytes()?);
    editor.changed = true;
    Ok(())
}

pub(crate) fn remove(editor: &mut Editor, id: u32) -> Result<ExternalObject> {
    let mut candidate = editor.clone();
    let removed = candidate.collection.remove(id)?;
    let persist = removed.persist_id();
    if !candidate
        .collection
        .objects
        .iter()
        .any(|object| object.persist_id() == persist)
    {
        candidate.removed_persist_ids.insert(persist);
        candidate.staged_storage.remove(&persist);
    }
    candidate.rewrite_object_list = true;
    candidate.changed = true;
    *editor = candidate;
    Ok(removed)
}

pub(crate) fn reorder(editor: &mut Editor, ids: &[u32]) -> Result<()> {
    let mut candidate = editor.clone();
    candidate.collection.reorder(ids)?;
    candidate.rewrite_object_list = true;
    candidate.changed = true;
    *editor = candidate;
    Ok(())
}

fn next_persist_id(editor: &Editor) -> Result<u32> {
    editor
        .mappings
        .keys()
        .chain(editor.staged_storage.keys())
        .copied()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .filter(|id| *id <= 0x000F_FFFF)
        .ok_or_else(|| Error::Corrupted("persist ID space exhausted".into()))
}

fn set_persist_id(object: &mut ExternalObject, persist_id: u32) {
    match object {
        ExternalObject::Object(value) => value.object.persist_id = persist_id,
        ExternalObject::ActiveXControl(value) => value.object.persist_id = persist_id,
    }
}
