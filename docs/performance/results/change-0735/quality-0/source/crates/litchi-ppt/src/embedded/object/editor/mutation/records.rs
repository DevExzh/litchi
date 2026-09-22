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

    // All typed checks are complete before this point.  The only remaining
    // operation is ownership transfer into the staging map followed by a
    // scalar flag update; cloning the complete editor provides no additional
    // recoverable failure boundary here.
    editor.staged_storage.insert(persist_id, record);
    editor.changed = true;
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

    // See the replacement path above: no fallible, typed operation follows
    // validation, so a full Editor clone is unnecessary.
    editor.staged_storage.insert(persist_id, record);
    editor.changed = true;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::super::{Collection, Editor};
    use super::{insert_persisted_record, replace_persisted_record};
    use crate::package::Error;
    use std::collections::{BTreeMap, HashSet};
    use std::sync::Arc;

    fn test_editor() -> Editor {
        Editor {
            original: Arc::from([1, 2, 3, 4]),
            max_output_bytes: 4096,
            streams: vec![
                (vec!["PowerPoint Document".into()], vec![10, 11, 12]),
                (vec!["Current User".into()], vec![20, 21, 22]),
                (vec!["Unrelated".into()], vec![30, 31, 32]),
            ],
            document_path: vec!["PowerPoint Document".into()],
            current_user_path: vec!["Current User".into()],
            document: vec![10, 11, 12],
            current_user: vec![20, 21, 22],
            mappings: BTreeMap::from([(7, 0), (11, 8)]),
            current_edit_offset: 64,
            document_persist_id: 7,
            collection: Collection {
                id_seed: 99,
                objects: Vec::new(),
                unknown_records: Vec::new(),
            },
            staged_storage: BTreeMap::from([(13, vec![40, 41]), (17, vec![50, 51, 52])]),
            removed_persist_ids: HashSet::new(),
            rewrite_object_list: false,
            changed: false,
            layout: litchi_cfb::SectorLayoutPolicy::Rewrite,
        }
    }

    fn malformed_record() -> Vec<u8> {
        let mut record = vec![0; 8];
        record[4..8].copy_from_slice(&1u32.to_le_bytes());
        record
    }

    fn trailing_record() -> Vec<u8> {
        let mut record = valid_record(80);
        record.push(81);
        record
    }

    fn valid_record(marker: u8) -> Vec<u8> {
        let payload = [marker, marker.wrapping_add(1)];
        let mut record = vec![0, 0, 0, 0];
        record.extend_from_slice(&(payload.len() as u32).to_le_bytes());
        record.extend_from_slice(&payload);
        record
    }

    fn oversized_record() -> Vec<u8> {
        vec![0; 128 * 1024 * 1024 + 1]
    }

    fn assert_same_state(actual: &Editor, expected: &Editor) {
        assert_eq!(actual.original.as_ref(), expected.original.as_ref());
        assert_eq!(actual.max_output_bytes, expected.max_output_bytes);
        assert_eq!(actual.streams, expected.streams);
        assert_eq!(actual.document_path, expected.document_path);
        assert_eq!(actual.current_user_path, expected.current_user_path);
        assert_eq!(actual.document, expected.document);
        assert_eq!(actual.current_user, expected.current_user);
        assert_eq!(actual.mappings, expected.mappings);
        assert_eq!(actual.current_edit_offset, expected.current_edit_offset);
        assert_eq!(actual.document_persist_id, expected.document_persist_id);
        assert_eq!(actual.collection, expected.collection);
        assert_eq!(actual.staged_storage, expected.staged_storage);
        assert_eq!(actual.removed_persist_ids, expected.removed_persist_ids);
        assert_eq!(actual.rewrite_object_list, expected.rewrite_object_list);
        assert_eq!(actual.changed, expected.changed);
        assert_eq!(actual.layout, expected.layout);
    }

    fn reference_replace(editor: &mut Editor, persist_id: u32, record: Vec<u8>) {
        let mut candidate = editor.clone();
        candidate.staged_storage.insert(persist_id, record);
        candidate.changed = true;
        *editor = candidate;
    }

    fn reference_insert(editor: &mut Editor, persist_id: u32, record: Vec<u8>) {
        let mut candidate = editor.clone();
        candidate.staged_storage.insert(persist_id, record);
        candidate.changed = true;
        *editor = candidate;
    }

    #[test]
    fn replacement_refusals_leave_every_editor_field_unchanged() {
        let cases = [
            (7_000, vec![0; 8]),
            (7, vec![0; 7]),
            (7, malformed_record()),
            (7, trailing_record()),
        ];
        for (persist_id, record) in cases {
            let before = test_editor();
            let mut actual = before.clone();
            let error = replace_persisted_record(&mut actual, persist_id, record).unwrap_err();
            assert!(matches!(error, Error::Corrupted(_)));
            assert_same_state(&actual, &before);
        }

        let mut before = test_editor();
        before.removed_persist_ids.insert(7);
        let expected = before.clone();
        let error = replace_persisted_record(&mut before, 7, vec![0; 8]).unwrap_err();
        assert!(matches!(error, Error::Corrupted(_)));
        assert_same_state(&before, &expected);
    }

    #[test]
    fn insertion_refusals_leave_every_editor_field_unchanged() {
        let cases = [
            (0, vec![0; 8]),
            (0x0010_0000, vec![0; 8]),
            (7, vec![0; 8]),
            (13, vec![0; 8]),
            (19, vec![0; 7]),
            (19, malformed_record()),
            (19, trailing_record()),
        ];
        for (persist_id, record) in cases {
            let before = test_editor();
            let mut actual = before.clone();
            let error = insert_persisted_record(&mut actual, persist_id, record).unwrap_err();
            assert!(matches!(error, Error::Corrupted(_)));
            assert_same_state(&actual, &before);
        }
    }

    #[test]
    fn oversized_records_are_refused_without_mutating_either_editor() {
        let before = test_editor();
        let mut replacement = before.clone();
        let error = replace_persisted_record(&mut replacement, 7, oversized_record()).unwrap_err();
        assert!(matches!(error, Error::Corrupted(_)));
        assert_same_state(&replacement, &before);

        let before = test_editor();
        let mut insertion = before.clone();
        let error = insert_persisted_record(&mut insertion, 19, oversized_record()).unwrap_err();
        assert!(matches!(error, Error::Corrupted(_)));
        assert_same_state(&insertion, &before);
    }

    #[test]
    fn replacement_and_insertion_match_the_previous_clone_state_transition() {
        let initial = test_editor();
        let mut direct = initial.clone();
        let mut reference = initial;

        replace_persisted_record(&mut direct, 7, valid_record(60)).unwrap();
        reference_replace(&mut reference, 7, valid_record(60));
        assert_same_state(&direct, &reference);

        insert_persisted_record(&mut direct, 19, valid_record(70)).unwrap();
        reference_insert(&mut reference, 19, valid_record(70));
        assert_same_state(&direct, &reference);
        assert_eq!(
            direct.staged_storage.get(&13).map(Vec::as_slice),
            Some(&[40, 41][..])
        );
        assert_eq!(
            direct.staged_storage.get(&17).map(Vec::as_slice),
            Some(&[50, 51, 52][..])
        );
    }

    #[test]
    fn replacement_output_bytes_match_the_previous_clone_transition() {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/poi/test-data/slideshow/45543.ppt"),
        )
        .unwrap();
        let initial = Editor::open_records(source).unwrap();
        let persist_id = initial.document_persist_id;
        let replacement = initial.persisted_record(persist_id).unwrap();

        let mut direct = initial.clone();
        let mut reference = initial;
        replace_persisted_record(&mut direct, persist_id, replacement.clone()).unwrap();
        reference_replace(&mut reference, persist_id, replacement);

        assert_eq!(direct.finish().unwrap(), reference.finish().unwrap());
    }
}
