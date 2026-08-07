use crate::archive::RawMessage;
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::wire::{
    append_length_delimited_field, parse_wire_fields, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};

use super::*;
use crate::text::TextStorageId;
use crate::text::bookmark_types::{TextBookmarkName, TextBookmarkSettings};
use crate::text::hyperlink_storage::{TABLE_ENTRIES_FIELD, locate_storage};
use prost::Message;

const BOOKMARK_TABLE_FIELD: u32 = 15;

fn fixture() -> (super::super::IWorkTextEditor, TextStorageId) {
    let pages = PagesEditor::create_with_text("Alpha Beta").unwrap();
    let storage_id = pages.body_storage().unwrap().id;
    (
        super::super::IWorkTextEditor::from_package(pages.into_package()),
        storage_id,
    )
}

fn range(start: usize, end: usize) -> TextRange {
    TextRange::from_utf16_indexes(start, end).unwrap()
}

fn named(value: &str) -> TextBookmarkSettings {
    TextBookmarkSettings::new().with_name(TextBookmarkName::new(value).unwrap())
}

fn storage_message_index(package: &crate::IWorkPackage, storage_id: TextStorageId) -> usize {
    locate_storage(package, storage_id.get(), BOOKMARK_TABLE)
        .unwrap()
        .message_index
}

#[test]
fn unknown_table_entry_and_bookmark_fields_survive_updates() {
    let (mut editor, storage_id) = fixture();
    let bookmark = editor
        .add_text_bookmark(storage_id, range(0, 5), named("Alpha"))
        .unwrap();
    let mut package = editor.into_package();
    let location = locate_storage(&package, storage_id.get(), BOOKMARK_TABLE).unwrap();
    let archive_name = location.archive_name.clone();
    let message_index = location.message_index;
    package
        .update_archive(&archive_name, |archive| {
            let storage = archive.object_mut(storage_id.get()).unwrap();
            let original = &storage.messages[message_index];
            let data =
                transform_length_delimited_field(&original.data, BOOKMARK_TABLE_FIELD, |table| {
                    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                    let first = patch_varint_field(entries[0], 77, false, Some(42))?;
                    let table = rewrite_repeated_length_delimited_fields(
                        table,
                        TABLE_ENTRIES_FIELD,
                        &[first, entries[1].to_vec()],
                    )?;
                    patch_varint_field(&table, 99, false, Some(7))
                })?;
            storage.replace_message(
                message_index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            let object = archive.object_mut(bookmark.id.get()).unwrap();
            let original = &object.messages[0];
            let data = patch_varint_field(&original.data, 88, false, Some(9))?;
            object.replace_message(
                0,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();

    let mut editor = super::super::IWorkTextEditor::from_package(package);
    editor
        .update_text_bookmark(storage_id, bookmark.id, range(6, 10), named("Beta"))
        .unwrap();
    let package = editor.into_package();
    let message_index = storage_message_index(&package, storage_id);
    let archive = package.archive(&archive_name).unwrap();
    let storage = archive.object(storage_id.get()).unwrap();
    let message = &storage.messages[message_index];
    let table = repeated_length_delimited_payloads(&message.data, BOOKMARK_TABLE_FIELD).unwrap()[0];
    assert!(
        parse_wire_fields(table)
            .unwrap()
            .iter()
            .any(|field| field.number() == 99)
    );
    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD).unwrap();
    assert!(
        parse_wire_fields(entries[1])
            .unwrap()
            .iter()
            .any(|field| field.number() == 77)
            || parse_wire_fields(entries[0])
                .unwrap()
                .iter()
                .any(|field| field.number() == 77)
    );
    let object = archive.object(bookmark.id.get()).unwrap();
    assert!(
        parse_wire_fields(&object.messages[0].data)
            .unwrap()
            .iter()
            .any(|field| field.number() == 88)
    );
}

#[test]
fn duplicate_bookmark_tables_fail_without_mutation() {
    let (mut editor, storage_id) = fixture();
    editor
        .add_text_bookmark(storage_id, range(0, 5), TextBookmarkSettings::new())
        .unwrap();
    let mut package = editor.into_package();
    let location = locate_storage(&package, storage_id.get(), BOOKMARK_TABLE).unwrap();
    let archive_name = location.archive_name.clone();
    let message_index = location.message_index;
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(storage_id.get()).unwrap();
            let original = &object.messages[message_index];
            let table =
                repeated_length_delimited_payloads(&original.data, BOOKMARK_TABLE_FIELD)?[0];
            let mut data = original.data.clone();
            append_length_delimited_field(&mut data, BOOKMARK_TABLE_FIELD, table)?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(editor.text_bookmarks(storage_id).is_err());
    assert!(
        editor
            .add_text_bookmark(storage_id, range(6, 10), named("Beta"))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn bookmark_mutations_use_the_resolved_storage_with_a_2022_style_sibling() {
    let (editor, storage_id) = fixture();
    let mut package = editor.into_package();
    let location = locate_storage(&package, storage_id.get(), BOOKMARK_TABLE).unwrap();
    let style_data = tswp::ParagraphStyleArchive {
        super_: crate::protobuf::tss::StyleArchive::default(),
        ..Default::default()
    }
    .encode_to_vec();
    package
        .update_archive(&location.archive_name, |archive| {
            archive
                .object_mut(storage_id.get())
                .unwrap()
                .push_message(RawMessage {
                    type_: 2_022,
                    data: style_data.clone(),
                })?;
            Ok(())
        })
        .unwrap();

    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let bookmark = editor
        .add_text_bookmark(storage_id, range(0, 5), named("Alpha"))
        .unwrap();
    editor
        .update_text_bookmark(storage_id, bookmark.id, range(6, 10), named("Beta"))
        .unwrap();
    editor
        .remove_text_bookmark(storage_id, bookmark.id)
        .unwrap();

    let package = editor.into_package();
    let location = locate_storage(&package, storage_id.get(), BOOKMARK_TABLE).unwrap();
    let archive = package.archive(&location.archive_name).unwrap();
    let object = archive.object(storage_id.get()).unwrap();
    assert_eq!(
        object.messages[location.message_index].type_,
        location.message_type
    );
    let sibling = object
        .messages
        .iter()
        .find(|message| message.type_ == 2_022 && message.data == style_data)
        .unwrap();
    assert_eq!(sibling.data, style_data);
}

#[test]
fn bookmark_with_an_additional_owner_cannot_be_updated_or_deleted() {
    let (mut editor, storage_id) = fixture();
    let bookmark = editor
        .add_text_bookmark(storage_id, range(0, 5), named("Alpha"))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id.get(), BOOKMARK_TABLE)
        .unwrap()
        .archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let other = archive
                .objects
                .iter_mut()
                .find(|object| {
                    object.archive_info.identifier != Some(storage_id.get())
                        && !object.archive_info.message_infos.is_empty()
                })
                .unwrap();
            other.archive_info.message_infos[0]
                .object_references
                .push(bookmark.id.get());
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_bookmark(storage_id, bookmark.id, range(6, 10), named("Beta"))
            .is_err()
    );
    assert!(
        editor
            .remove_text_bookmark(storage_id, bookmark.id)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
