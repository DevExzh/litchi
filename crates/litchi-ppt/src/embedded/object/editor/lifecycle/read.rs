//! Read-only views over the editor's live persisted-record state.

use super::super::{Editor, Result, rewrite};
use crate::package::Error;

pub(crate) fn persist_ids(editor: &Editor) -> Vec<u32> {
    editor.mappings.keys().copied().collect()
}

pub(crate) fn persisted_record(editor: &Editor, persist_id: u32) -> Result<Vec<u8>> {
    if let Some(record) = editor.staged_storage.get(&persist_id) {
        return Ok(record.clone());
    }
    let offset = *editor
        .mappings
        .get(&persist_id)
        .ok_or_else(|| Error::Corrupted(format!("unknown persist ID {persist_id}")))?
        as usize;
    Ok(rewrite::slice(&editor.document, offset)?.to_vec())
}
