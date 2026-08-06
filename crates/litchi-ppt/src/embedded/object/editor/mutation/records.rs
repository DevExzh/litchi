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
        || rewrite::slice(&record, 0).map(|value| value.len())? != record.len()
    {
        return Err(Error::Corrupted(
            "replacement persisted record has an invalid length".into(),
        ));
    }

    let mut candidate = editor.clone();
    candidate.staged_storage.insert(persist_id, record);
    candidate.changed = true;
    *editor = candidate;
    Ok(())
}
