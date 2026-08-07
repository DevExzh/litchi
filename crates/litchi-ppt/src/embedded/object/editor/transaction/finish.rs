//! Emits one append-only PPT UserEdit transaction.

use super::super::{Editor, Result, mapping, rewrite};
use crate::package::Error;
use crate::writer::{PersistPtrBuilder, UserEditAtom};
use litchi_cfb::{OleFile, OleWriter};
use std::collections::BTreeMap;
use std::io::{self, Cursor, Seek, SeekFrom, Write};

pub(crate) fn finish(mut editor: Editor) -> Result<Vec<u8>> {
    if !editor.changed {
        ensure_output_limit(editor.original.len(), editor.max_output_bytes, "source")?;
        let mut original = Vec::new();
        original
            .try_reserve_exact(editor.original.len())
            .map_err(|_| Error::AllocationFailed("PowerPoint editor source"))?;
        original.extend_from_slice(&editor.original);
        return Ok(original);
    }

    for id in &editor.removed_persist_ids {
        editor.mappings.remove(id);
    }

    let mut projected_len = editor.document.len();
    ensure_output_limit(
        projected_len,
        editor.max_output_bytes,
        "incremental document stream",
    )?;
    for (id, record) in &editor.staged_storage {
        editor.mappings.insert(
            *id,
            u32::try_from(projected_len)
                .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?,
        );
        projected_len =
            checked_projected_len(projected_len, record.len(), editor.max_output_bytes)?;
    }

    let rewritten_document = if editor.rewrite_object_list {
        let document = rewritten_object_list(&editor)?;
        editor.mappings.insert(
            editor.document_persist_id,
            u32::try_from(projected_len)
                .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?,
        );
        projected_len =
            checked_projected_len(projected_len, document.len(), editor.max_output_bytes)?;
        Some(document)
    } else {
        None
    };

    let persist_dir_offset = u32::try_from(projected_len)
        .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?;
    projected_len = checked_projected_len(
        projected_len,
        persist_directory_record_len(&editor.mappings)?,
        editor.max_output_bytes,
    )?;
    let new_edit_offset = u32::try_from(projected_len)
        .map_err(|_| Error::Corrupted("PPT stream exceeds u32".into()))?;
    projected_len = checked_projected_len(projected_len, 36, editor.max_output_bytes)?;

    let mut appended = Vec::new();
    appended
        .try_reserve_exact(projected_len)
        .map_err(|_| Error::AllocationFailed("PowerPoint incremental document stream"))?;
    append_checked(&mut appended, &editor.document, editor.max_output_bytes)?;
    for record in editor.staged_storage.values() {
        append_checked(&mut appended, record, editor.max_output_bytes)?;
    }
    if let Some(document) = &rewritten_document {
        append_checked(&mut appended, document, editor.max_output_bytes)?;
    }

    let mut builder = PersistPtrBuilder::new();
    for (id, offset) in &editor.mappings {
        builder.set_offset(*id, *offset);
    }
    append_checked(
        &mut appended,
        &builder.generate_full_record(),
        editor.max_output_bytes,
    )?;

    let max_id = editor
        .mappings
        .keys()
        .next_back()
        .copied()
        .unwrap_or(editor.document_persist_id);
    let mut edit =
        UserEditAtom::new_minimal(persist_dir_offset, editor.document_persist_id, max_id, 0);
    edit.offset_last_edit = editor.current_edit_offset;
    append_checked(
        &mut appended,
        &edit.generate_record(),
        editor.max_output_bytes,
    )?;
    editor.current_user[16..20].copy_from_slice(&new_edit_offset.to_le_bytes());

    let bytes = write_package(&editor, &appended)?;
    validate_rewrite(&editor, bytes)
}

fn rewritten_object_list(editor: &Editor) -> Result<Vec<u8>> {
    let document_offset = *editor
        .mappings
        .get(&editor.document_persist_id)
        .ok_or_else(|| Error::Corrupted("Document persist mapping is missing".into()))?
        as usize;
    let old_document = rewrite::slice(&editor.document, document_offset)?;
    rewrite::replace_nested_record(
        old_document,
        rewrite::external_object_list(),
        &editor.collection.to_record_bytes()?,
    )
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
    let mut output = BoundedCursor::new(editor.max_output_bytes);
    if let Err(error) = writer.write_to(&mut output) {
        if output.limit_exceeded() {
            return Err(output_limit_error(editor.max_output_bytes, "OLE package"));
        }
        return Err(error.into());
    }
    Ok(output.into_inner())
}

fn validate_rewrite(editor: &Editor, bytes: Vec<u8>) -> Result<Vec<u8>> {
    let mut reopen = OleFile::open(Cursor::new(bytes.as_slice()))?;
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

fn checked_projected_len(current: usize, additional: usize, maximum: usize) -> Result<usize> {
    let projected = current.checked_add(additional).ok_or_else(|| {
        Error::ResourceLimit("PowerPoint incremental document stream size overflows".into())
    })?;
    ensure_output_limit(projected, maximum, "incremental document stream")?;
    Ok(projected)
}

fn ensure_output_limit(size: usize, maximum: usize, context: &str) -> Result<()> {
    if size > maximum {
        return Err(Error::ResourceLimit(format!(
            "PowerPoint editor {context} requires {size} bytes, exceeding the {maximum}-byte output limit"
        )));
    }
    Ok(())
}

fn output_limit_error(maximum: usize, context: &str) -> Error {
    Error::ResourceLimit(format!(
        "PowerPoint editor {context} exceeds the {maximum}-byte output limit"
    ))
}

fn append_checked(output: &mut Vec<u8>, bytes: &[u8], maximum: usize) -> Result<()> {
    let projected = checked_projected_len(output.len(), bytes.len(), maximum)?;
    output
        .try_reserve(bytes.len())
        .map_err(|_| Error::AllocationFailed("PowerPoint incremental document stream"))?;
    output.extend_from_slice(bytes);
    debug_assert_eq!(output.len(), projected);
    Ok(())
}

fn persist_directory_record_len(mappings: &BTreeMap<u32, u32>) -> Result<usize> {
    let mut runs = 0usize;
    let mut prior: Option<u32> = None;
    for id in mappings.keys().copied() {
        if prior.is_none_or(|value| value.checked_add(1) != Some(id)) {
            runs = runs.checked_add(1).ok_or_else(|| {
                Error::ResourceLimit("PowerPoint persist directory size overflows".into())
            })?;
        }
        prior = Some(id);
    }
    mappings
        .len()
        .checked_add(runs)
        .and_then(|words| words.checked_mul(4))
        .and_then(|payload| payload.checked_add(8))
        .ok_or_else(|| Error::ResourceLimit("PowerPoint persist directory size overflows".into()))
}

struct BoundedCursor {
    inner: Cursor<Vec<u8>>,
    maximum: usize,
    limit_exceeded: bool,
}

impl BoundedCursor {
    fn new(maximum: usize) -> Self {
        Self {
            inner: Cursor::new(Vec::new()),
            maximum,
            limit_exceeded: false,
        }
    }

    fn limit_exceeded(&self) -> bool {
        self.limit_exceeded
    }

    fn into_inner(self) -> Vec<u8> {
        self.inner.into_inner()
    }

    fn limit_error(&mut self) -> io::Error {
        self.limit_exceeded = true;
        io::Error::other("PowerPoint editor output limit exceeded")
    }
}

impl Write for BoundedCursor {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let position = usize::try_from(self.inner.position()).map_err(|_| self.limit_error())?;
        let end = position
            .checked_add(bytes.len())
            .ok_or_else(|| self.limit_error())?;
        if end > self.maximum {
            return Err(self.limit_error());
        }
        let additional = end.saturating_sub(self.inner.get_ref().len());
        self.inner
            .get_mut()
            .try_reserve(additional)
            .map_err(|_| io::Error::other("PowerPoint editor output allocation failed"))?;
        self.inner.write(bytes)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.inner.flush()
    }
}

impl Seek for BoundedCursor {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let previous = self.inner.position();
        let next = self.inner.seek(position)?;
        let maximum = u64::try_from(self.maximum).unwrap_or(u64::MAX);
        if next > maximum {
            self.inner.set_position(previous);
            return Err(self.limit_error());
        }
        Ok(next)
    }
}

#[cfg(test)]
mod bounded_cursor_tests {
    use super::BoundedCursor;
    use std::io::{Seek, SeekFrom, Write};

    #[test]
    fn rejects_writes_and_seeks_beyond_the_output_limit() {
        let mut cursor = BoundedCursor::new(4);
        cursor.write_all(b"four").unwrap();
        assert!(cursor.write_all(b"!").is_err());
        assert!(cursor.seek(SeekFrom::Start(5)).is_err());
        assert_eq!(cursor.into_inner(), b"four");
    }
}
