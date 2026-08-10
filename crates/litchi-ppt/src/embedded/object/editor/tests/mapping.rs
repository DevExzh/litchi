use super::super::{Editor, mapping, rewrite};
use super::ppt_record;
use crate::writer::{PersistPtrBuilder, UserEditAtom};
use litchi_cfb::OleFile;
use std::io::Cursor;

#[test]
fn merges_newest_incremental_mapping_over_prior_edit() {
    let object1 = ppt_record(0, 0x1111, b"one");
    let object2 = ppt_record(0, 0x2222, b"two");
    let mut document = object1.clone();
    document.extend_from_slice(&object2);
    let mut first_dir = PersistPtrBuilder::new();
    first_dir.set_offset(1, 0);
    let first_dir_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&first_dir.generate_full_record());
    let first_edit_offset = u32::try_from(document.len()).unwrap();
    document
        .extend_from_slice(&UserEditAtom::new_minimal(first_dir_offset, 1, 1, 0).generate_record());
    let replacement_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&object2);
    let mut second_dir = PersistPtrBuilder::new();
    second_dir.set_offset(1, replacement_offset);
    let second_dir_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&second_dir.generate_incremental_record());
    let mut edit = UserEditAtom::new_minimal(second_dir_offset, 1, 1, 0);
    edit.offset_last_edit = first_edit_offset;
    let second_edit_offset = u32::try_from(document.len()).unwrap();
    document.extend_from_slice(&edit.generate_record());
    let (mapping, document_id) = mapping::read(&document, second_edit_offset).unwrap();
    assert_eq!(document_id, 1);
    assert_eq!(mapping.get(&1), Some(&replacement_offset));
    assert_eq!(rewrite::type_of(&document[..8]).unwrap(), 0x1111);
}

#[test]
fn editor_publishes_native_incremental_directory_and_retains_prior_chain() {
    let source = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/slideshow/45543.ppt"),
    )
    .unwrap();
    let mut editor = Editor::open_records(source).unwrap();
    let previous_edit_offset = editor.current_edit_offset;
    let document_persist_id = editor.document_persist_id;
    let replacement = editor.persisted_record(document_persist_id).unwrap();
    editor
        .replace_persisted_record(document_persist_id, replacement.clone())
        .unwrap();
    let published = editor.finish().unwrap();

    let mut ole = OleFile::open(Cursor::new(published.as_slice())).unwrap();
    let document = ole.open_stream(&["PowerPoint Document"]).unwrap();
    let current_user = ole.open_stream(&["Current User"]).unwrap();
    let latest_edit_offset = rewrite::u32_at(&current_user, 16).unwrap();
    assert_ne!(latest_edit_offset, previous_edit_offset);
    let latest_edit = rewrite::slice(&document, latest_edit_offset as usize).unwrap();
    assert_eq!(rewrite::type_of(latest_edit).unwrap(), 4085);
    assert_eq!(
        rewrite::u32_at(&latest_edit[8..], 8).unwrap(),
        previous_edit_offset
    );
    let latest_directory_offset = rewrite::u32_at(&latest_edit[8..], 12).unwrap();
    let latest_directory = rewrite::slice(&document, latest_directory_offset as usize).unwrap();
    assert_eq!(
        rewrite::type_of(latest_directory).unwrap(),
        6002,
        "native PPT importers require PersistPtrIncrementalBlock in a UserEdit chain"
    );

    let (prior_mapping, prior_document_id) =
        mapping::read(&document, previous_edit_offset).unwrap();
    assert_eq!(prior_document_id, document_persist_id);
    assert!(prior_mapping.contains_key(&document_persist_id));
    let (latest_mapping, latest_document_id) =
        mapping::read(&document, latest_edit_offset).unwrap();
    assert_eq!(latest_document_id, document_persist_id);
    assert!(
        latest_mapping[&document_persist_id] > prior_mapping[&document_persist_id],
        "latest directory must resolve the appended replacement"
    );

    let reopened = Editor::open_records(published).unwrap();
    assert_eq!(
        reopened.persisted_record(document_persist_id).unwrap(),
        replacement
    );
}
