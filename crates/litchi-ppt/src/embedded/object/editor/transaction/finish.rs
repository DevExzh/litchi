//! Emits one append-only PPT `UserEdit` transaction.

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
            .map_err(|_err| Error::AllocationFailed("PowerPoint editor source"))?;
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
                .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?,
        );
        projected_len =
            checked_projected_len(projected_len, record.len(), editor.max_output_bytes)?;
    }

    let rewritten_document = if editor.rewrite_object_list {
        let document = rewritten_object_list(&editor)?;
        editor.mappings.insert(
            editor.document_persist_id,
            u32::try_from(projected_len)
                .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?,
        );
        projected_len =
            checked_projected_len(projected_len, document.len(), editor.max_output_bytes)?;
        Some(document)
    } else {
        None
    };

    let persist_dir_offset = u32::try_from(projected_len)
        .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?;
    projected_len = checked_projected_len(
        projected_len,
        persist_directory_record_len(&editor.mappings)?,
        editor.max_output_bytes,
    )?;
    let new_edit_offset = u32::try_from(projected_len)
        .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?;
    projected_len = checked_projected_len(projected_len, 36, editor.max_output_bytes)?;

    let mut appended = Vec::new();
    appended
        .try_reserve_exact(projected_len)
        .map_err(|_err| Error::AllocationFailed("PowerPoint incremental document stream"))?;
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
        &builder.generate_incremental_record(),
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

    let bytes = write_package(&mut editor, appended)?;
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

fn write_package(editor: &mut Editor, appended: Vec<u8>) -> Result<Vec<u8>> {
    let mut writer = OleWriter::new();
    writer.set_sector_layout_policy(editor.layout);
    writer.adopt_source_layout(&editor.original)?;

    // `find_stream` selects the first matching leaf name.  A valid CFB
    // directory cannot contain the same complete path twice, but a malformed
    // or synthetic Editor can still carry duplicate complete paths.  Keep the
    // old borrowed-copy behavior for those paths so the ownership handoff does
    // not narrow the reachable error or replacement semantics.
    let document_matches = editor
        .streams
        .iter()
        .filter(|(path, _data)| path == &editor.document_path)
        .count();
    let current_user_matches = editor
        .streams
        .iter()
        .filter(|(path, _data)| path == &editor.current_user_path)
        .count();
    let mut appended = Some(appended);
    let current_user = std::mem::take(&mut editor.current_user);
    let mut current_user = Some(current_user);
    let streams = std::mem::take(&mut editor.streams);

    for (path, data) in streams {
        let path_refs = stream_refs(&path);
        if path == editor.document_path {
            if document_matches == 1 {
                let payload = appended.take().ok_or_else(|| {
                    Error::Corrupted("PPT Document stream payload was consumed twice".into())
                })?;
                writer.create_stream_owned(&path_refs, payload)?;
            } else {
                let payload = appended.as_deref().ok_or_else(|| {
                    Error::Corrupted("PPT Document stream payload is missing".into())
                })?;
                writer.create_stream(&path_refs, payload)?;
            }
        } else if path == editor.current_user_path {
            if current_user_matches == 1 {
                let payload = current_user.take().ok_or_else(|| {
                    Error::Corrupted("PPT Current User stream payload was consumed twice".into())
                })?;
                writer.create_stream_owned(&path_refs, payload)?;
            } else {
                let payload = current_user.as_deref().ok_or_else(|| {
                    Error::Corrupted("PPT Current User stream payload is missing".into())
                })?;
                writer.create_stream(&path_refs, payload)?;
            }
        } else {
            writer.create_stream_owned(&path_refs, data)?;
        }
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
        .map_err(|_err| Error::AllocationFailed("PowerPoint incremental document stream"))?;
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

#[allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the bounded cursor is grouped directly with its trait implementations"
)]
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
        let position = usize::try_from(self.inner.position()).map_err(|_err| self.limit_error())?;
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
            .map_err(|_err| io::Error::other("PowerPoint editor output allocation failed"))?;
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
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::too_many_lines,
    reason = "these tests compare the complete old and candidate finish paths"
)]
mod candidate_finish_tests {
    use super::super::super::{Editor, Result};
    use super::{
        BoundedCursor, append_checked, checked_projected_len, output_limit_error,
        persist_directory_record_len, stream_refs, validate_rewrite,
    };
    use crate::package::Error;
    use crate::writer::{PersistPtrBuilder, UserEditAtom};
    use litchi_cfb::{OleFile, OleWriter, SectorLayoutPolicy};
    use std::io::Cursor;
    use std::sync::Arc;

    fn fixture() -> Vec<u8> {
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/slideshow/45543.ppt"),
        )
        .expect("PPT fixture")
    }

    /// This is deliberately test-only: it is the pre-0734 writer ingress used
    /// as a byte-parity oracle for the ownership handoff.
    fn write_package_reference(editor: &Editor, appended: &[u8]) -> Result<Vec<u8>> {
        let mut writer = OleWriter::new();
        writer.set_sector_layout_policy(editor.layout);
        writer.adopt_source_layout(&editor.original)?;
        for (path, data) in &editor.streams {
            let stream_data = if path == &editor.document_path {
                appended
            } else if path == &editor.current_user_path {
                &editor.current_user
            } else {
                data
            };
            writer.create_stream(&stream_refs(path), stream_data)?;
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

    fn finish_reference(mut editor: Editor) -> Result<Vec<u8>> {
        assert!(
            editor.changed,
            "the reference helper covers changed commits"
        );

        for id in &editor.removed_persist_ids {
            editor.mappings.remove(id);
        }

        let mut projected_len = editor.document.len();
        super::ensure_output_limit(
            projected_len,
            editor.max_output_bytes,
            "incremental document stream",
        )?;
        for (id, record) in &editor.staged_storage {
            editor.mappings.insert(
                *id,
                u32::try_from(projected_len)
                    .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?,
            );
            projected_len =
                checked_projected_len(projected_len, record.len(), editor.max_output_bytes)?;
        }

        let rewritten_document = if editor.rewrite_object_list {
            let document = super::rewritten_object_list(&editor)?;
            editor.mappings.insert(
                editor.document_persist_id,
                u32::try_from(projected_len)
                    .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?,
            );
            projected_len =
                checked_projected_len(projected_len, document.len(), editor.max_output_bytes)?;
            Some(document)
        } else {
            None
        };

        let persist_dir_offset = u32::try_from(projected_len)
            .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?;
        projected_len = checked_projected_len(
            projected_len,
            persist_directory_record_len(&editor.mappings)?,
            editor.max_output_bytes,
        )?;
        let new_edit_offset = u32::try_from(projected_len)
            .map_err(|_err| Error::Corrupted("PPT stream exceeds u32".into()))?;
        projected_len = checked_projected_len(projected_len, 36, editor.max_output_bytes)?;

        let mut appended = Vec::new();
        appended
            .try_reserve_exact(projected_len)
            .map_err(|_err| Error::AllocationFailed("PowerPoint incremental document stream"))?;
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
            &builder.generate_incremental_record(),
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

        let bytes = write_package_reference(&editor, &appended)?;
        validate_rewrite(&editor, bytes)
    }

    #[test]
    fn owned_writer_matches_borrowed_writer_for_reuse_and_rewrite() {
        let source = fixture();
        for policy in [SectorLayoutPolicy::Reuse, SectorLayoutPolicy::Rewrite] {
            let mut candidate = Editor::open_records(source.clone()).unwrap();
            candidate.set_sector_layout_policy(policy);
            let reference = candidate.clone();
            let appended = candidate.document.clone();
            let expected = write_package_reference(&reference, &appended).unwrap();
            let actual = super::write_package(&mut candidate, appended).unwrap();
            assert_eq!(actual, expected, "policy {policy:?}");
        }
    }

    #[test]
    fn changed_record_finish_matches_borrowed_finish_and_reopens() {
        let source = fixture();
        let persist_id = Editor::open_records(source.clone())
            .unwrap()
            .document_persist_id;
        for policy in [SectorLayoutPolicy::Reuse, SectorLayoutPolicy::Rewrite] {
            let mut candidate = Editor::open_records(source.clone()).unwrap();
            let replacement = candidate.persisted_record(persist_id).unwrap();
            candidate
                .replace_persisted_record(persist_id, replacement.clone())
                .unwrap();
            candidate.set_sector_layout_policy(policy);

            let mut reference = Editor::open_records(source.clone()).unwrap();
            reference
                .replace_persisted_record(persist_id, replacement.clone())
                .unwrap();
            reference.set_sector_layout_policy(policy);

            let expected = finish_reference(reference).unwrap();
            let actual = candidate.finish().unwrap();
            assert_eq!(actual, expected, "policy {policy:?}");
            let reopened = Editor::open_records(actual).unwrap();
            assert_eq!(
                reopened.persisted_record(persist_id).unwrap(),
                replacement,
                "policy {policy:?} replacement record"
            );
        }
    }

    #[test]
    fn exact_noop_and_changed_finish_leave_the_source_immutable() {
        let source = fixture();
        let source_arc: Arc<[u8]> = Arc::from(source.clone());
        let noop = Editor::open_records_arc_with_limit(source_arc.clone(), usize::MAX)
            .unwrap()
            .finish()
            .unwrap();
        assert_eq!(noop, source);
        assert_eq!(source_arc.as_ref(), source.as_slice());

        let mut changed =
            Editor::open_records_arc_with_limit(source_arc.clone(), usize::MAX).unwrap();
        let persist_id = changed.document_persist_id;
        let record = changed.persisted_record(persist_id).unwrap();
        changed
            .replace_persisted_record(persist_id, record)
            .unwrap();
        changed.finish().unwrap();
        assert_eq!(source_arc.as_ref(), source.as_slice());
    }

    #[test]
    fn projected_output_limit_remains_typed() {
        let source = fixture();
        let mut editor = Editor::open_records(source).unwrap();
        editor.max_output_bytes = editor.document.len();
        let persist_id = editor.document_persist_id;
        let record = editor.persisted_record(persist_id).unwrap();
        editor.replace_persisted_record(persist_id, record).unwrap();
        let error = editor.finish().unwrap_err();
        assert!(matches!(
            error,
            Error::ResourceLimit(message) if message.contains("incremental document stream")
        ));
    }

    #[test]
    fn duplicate_selected_paths_keep_the_borrowed_fallback() {
        let source = fixture();
        let mut candidate = Editor::open_records(source).unwrap();
        let duplicate_document = (candidate.document_path.clone(), candidate.document.clone());
        let duplicate_current_user = (
            candidate.current_user_path.clone(),
            candidate.current_user.clone(),
        );
        candidate.streams.push(duplicate_document);
        candidate.streams.push(duplicate_current_user);
        let reference = candidate.clone();
        let appended = candidate.document.clone();
        let expected = write_package_reference(&reference, &appended).unwrap();
        let actual = super::write_package(&mut candidate, appended).unwrap();
        assert_eq!(actual, expected);

        let mut ole = OleFile::open(Cursor::new(actual)).unwrap();
        assert!(
            !ole.open_stream(&stream_refs(&reference.document_path))
                .unwrap()
                .is_empty()
        );
    }
}
