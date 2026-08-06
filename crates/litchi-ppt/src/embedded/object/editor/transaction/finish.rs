//! Emits one append-only PPT UserEdit transaction.

use super::super::{Editor, Result, mapping, rewrite};
use crate::package::Error;
use crate::writer::{PersistPtrBuilder, UserEditAtom};
use litchi_cfb::{OleFile, OleWriter};
use std::io::Cursor;

pub(crate) fn finish(mut editor: Editor) -> Result<Vec<u8>> {
    if !editor.changed {
        return Ok(editor.original);
    }

    for id in &editor.removed_persist_ids {
        editor.mappings.remove(id);
    }

    let mut appended = editor.document.clone();
    for (id, record) in &editor.staged_storage {
        editor.mappings.insert(
            *id,
            u32::try_from(appended.len())
                .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?,
        );
        appended.extend_from_slice(record);
    }

    if editor.rewrite_object_list {
        append_rewritten_object_list(&mut editor, &mut appended)?;
    }

    let persist_dir_offset = u32::try_from(appended.len())
        .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?;
    let mut builder = PersistPtrBuilder::new();
    for (id, offset) in &editor.mappings {
        builder.set_offset(*id, *offset);
    }
    appended.extend_from_slice(&builder.generate_full_record());

    let max_id = editor
        .mappings
        .keys()
        .next_back()
        .copied()
        .unwrap_or(editor.document_persist_id);
    let mut edit =
        UserEditAtom::new_minimal(persist_dir_offset, editor.document_persist_id, max_id, 0);
    edit.offset_last_edit = editor.current_edit_offset;
    let new_edit_offset = u32::try_from(appended.len())
        .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?;
    appended.extend_from_slice(&edit.generate_record());
    editor.current_user[16..20].copy_from_slice(&new_edit_offset.to_le_bytes());

    let bytes = write_package(&editor, &appended)?;
    validate_rewrite(&editor, bytes)
}

fn append_rewritten_object_list(editor: &mut Editor, appended: &mut Vec<u8>) -> Result<()> {
    let document_offset = *editor
        .mappings
        .get(&editor.document_persist_id)
        .ok_or_else(|| Error::Corrupted("Document persist mapping is missing".into()))?
        as usize;
    let old_document = rewrite::slice(&editor.document, document_offset)?;
    let new_document = rewrite::replace_nested_record(
        old_document,
        rewrite::external_object_list(),
        &editor.collection.to_record_bytes()?,
    )?;
    editor.mappings.insert(
        editor.document_persist_id,
        u32::try_from(appended.len())
            .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?,
    );
    appended.extend_from_slice(&new_document);
    Ok(())
}

fn write_package(editor: &Editor, appended: &[u8]) -> Result<Vec<u8>> {
    let mut writer = OleWriter::new();
    for (path, data) in &editor.streams {
        let data = if path == &editor.document_path {
            appended
        } else if path == &editor.current_user_path {
            &editor.current_user
        } else {
            data
        };
        writer.create_stream(&stream_refs(path), data)?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn validate_rewrite(editor: &Editor, bytes: Vec<u8>) -> Result<Vec<u8>> {
    let mut reopen = OleFile::open(Cursor::new(bytes.clone()))?;
    let document = reopen.open_stream(&stream_refs(&editor.document_path))?;
    let current_user = reopen.open_stream(&stream_refs(&editor.current_user_path))?;
    let (mapping, _) = mapping::read(&document, rewrite::u32_at(&current_user, 16)?)?;
    for object in &editor.collection.objects {
        if !mapping.contains_key(&object.persist_id()) {
            return Err(Error::Corrupted(
                "rewritten persist mapping failed validation".into(),
            ));
        }
    }
    Ok(bytes)
}

fn stream_refs(path: &[String]) -> Vec<&str> {
    path.iter().map(String::as_str).collect()
}
