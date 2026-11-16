//! Staged complete-record replacement.

use super::super::{Editor, Result, rewrite};
use crate::package::Error;

pub(crate) fn replace_persisted_record(
    editor: &mut Editor,
    persist_id: u32,
    record: Vec<u8>,
) -> Result<()> {
    if !editor.mappings.contains_key(&persist_id)
        || editor.removed_persist_ids.contains(&persist_id)
    {
        return Err(Error::Corrupted(format!("unknown persist ID {persist_id}")));
    }
    if record.len() < 8
        || record.len() > 128 * 1024 * 1024
        || rewrite::slice(&record, 0).map(<[u8]>::len)? != record.len()
    {
        return Err(Error::Corrupted(
            "replacement persisted record has an invalid length".into(),
        ));
    }

    let mut candidate = editor.clone();
    candidate.staged_storage.insert(persist_id, record);
    candidate.state = candidate.state.with_staged_change();
    *editor = candidate;
    Ok(())
}

pub(crate) fn insert_persisted_record(
    editor: &mut Editor,
    persist_id: u32,
    record: Vec<u8>,
) -> Result<()> {
    if persist_id == 0
        || persist_id > 0x000f_ffff
        || editor.mappings.contains_key(&persist_id)
        || editor.staged_storage.contains_key(&persist_id)
    {
        return Err(Error::Corrupted(format!(
            "persist ID {persist_id} is not available for insertion"
        )));
    }
    if record.len() < 8
        || record.len() > 128 * 1024 * 1024
        || rewrite::slice(&record, 0).map(<[u8]>::len)? != record.len()
    {
        return Err(Error::Corrupted(
            "inserted persisted record has an invalid length".into(),
        ));
    }

    let mut candidate = editor.clone();
    candidate.staged_storage.insert(persist_id, record);
    candidate.state = candidate.state.with_staged_change();
    *editor = candidate;
    Ok(())
}

/// Stage the removal of a live persisted record, or drop a staged insertion.
#[cfg(feature = "equation")]
pub(crate) fn remove_persisted_record(editor: &mut Editor, persist_id: u32) -> Result<()> {
    if persist_id == editor.document_persist_id {
        return Err(Error::Corrupted(
            "the live document record cannot be removed".into(),
        ));
    }
    let live = editor.mappings.contains_key(&persist_id);
    if !live && !editor.staged_storage.contains_key(&persist_id) {
        return Err(Error::Corrupted(format!("unknown persist ID {persist_id}")));
    }
    let mut candidate = editor.clone();
    candidate.staged_storage.remove(&persist_id);
    if live {
        candidate.removed_persist_ids.insert(persist_id);
    }
    candidate.state = candidate.state.with_staged_change();
    *editor = candidate;
    Ok(())
}
