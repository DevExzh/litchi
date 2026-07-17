use crate::archive::{ArchiveObject, RawMessage};
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::wire::{
    append_length_delimited_field, parse_wire_fields, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};

use super::*;
use crate::text::hyperlink_storage::{SMART_FIELD_TABLE_FIELD, TABLE_ENTRIES_FIELD};
use crate::text::storage_wire::STORAGE_MESSAGE_TYPES;

fn fixture() -> (super::super::IWorkTextEditor, u64) {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages
        .add_text_box(
            4,
            "Alpha Beta",
            DrawablePoint { x: 40.0, y: 80.0 },
            DrawableSize {
                width: 320.0,
                height: 140.0,
            },
        )
        .unwrap();
    (
        super::super::IWorkTextEditor::from_package(pages.into_package()),
        text_box.storage.object_id,
    )
}

fn range(start: usize, end: usize) -> TextRange {
    TextRange::from_utf16_indexes(start, end).unwrap()
}

fn target(value: &str) -> TextHyperlinkTarget {
    TextHyperlinkTarget::new(value).unwrap()
}

fn storage_message_index(object: &ArchiveObject) -> usize {
    object
        .messages
        .iter()
        .position(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .unwrap()
}

#[test]
fn unknown_table_entry_and_hyperlink_fields_survive_updates() {
    let (mut editor, storage_id) = fixture();
    let hyperlink = editor
        .add_text_hyperlink(storage_id, range(0, 5), target("https://example.com/one"))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let storage = archive.object_mut(storage_id).unwrap();
            let index = storage_message_index(storage);
            let original = &storage.messages[index];
            let data = transform_length_delimited_field(
                &original.data,
                SMART_FIELD_TABLE_FIELD,
                |table| {
                    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                    let first = patch_varint_field(entries[0], 77, false, Some(42))?;
                    let table = rewrite_repeated_length_delimited_fields(
                        table,
                        TABLE_ENTRIES_FIELD,
                        &[first, entries[1].to_vec()],
                    )?;
                    patch_varint_field(&table, 99, false, Some(7))
                },
            )?;
            storage.replace_message(
                index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            let object = archive.object_mut(hyperlink.id.object_id()).unwrap();
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
        .update_text_hyperlink(
            storage_id,
            hyperlink.id,
            range(0, 5),
            target("mailto:two@example.com"),
        )
        .unwrap();
    let package = editor.into_package();
    let archive = package.archive(&archive_name).unwrap();
    let storage = archive.object(storage_id).unwrap();
    let message = &storage.messages[storage_message_index(storage)];
    let table =
        repeated_length_delimited_payloads(&message.data, SMART_FIELD_TABLE_FIELD).unwrap()[0];
    assert!(
        parse_wire_fields(table)
            .unwrap()
            .iter()
            .any(|field| field.number == 99)
    );
    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD).unwrap();
    assert!(
        parse_wire_fields(entries[0])
            .unwrap()
            .iter()
            .any(|field| field.number == 77)
    );
    let object = archive.object(hyperlink.id.object_id()).unwrap();
    assert!(
        parse_wire_fields(&object.messages[0].data)
            .unwrap()
            .iter()
            .any(|field| field.number == 88)
    );
}

#[test]
fn duplicate_smart_field_tables_fail_without_mutation() {
    let (mut editor, storage_id) = fixture();
    editor
        .add_text_hyperlink(storage_id, range(0, 5), target("https://example.com"))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(storage_id).unwrap();
            let index = storage_message_index(object);
            let original = &object.messages[index];
            let table =
                repeated_length_delimited_payloads(&original.data, SMART_FIELD_TABLE_FIELD)?[0];
            let mut data = original.data.clone();
            append_length_delimited_field(&mut data, SMART_FIELD_TABLE_FIELD, table)?;
            object.replace_message(
                index,
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
    assert!(editor.text_hyperlinks(storage_id).is_err());
    assert!(
        editor
            .add_text_hyperlink(storage_id, range(6, 10), target("?slide=next"))
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn hyperlink_with_an_additional_owner_cannot_be_updated_or_deleted() {
    let (mut editor, storage_id) = fixture();
    let hyperlink = editor
        .add_text_hyperlink(storage_id, range(0, 5), target("https://example.com"))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let other = archive
                .objects
                .iter_mut()
                .find(|object| {
                    object.archive_info.identifier != Some(storage_id)
                        && !object.archive_info.message_infos.is_empty()
                })
                .unwrap();
            other.archive_info.message_infos[0]
                .object_references
                .push(hyperlink.id.object_id());
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_hyperlink(
                storage_id,
                hyperlink.id,
                range(6, 10),
                target("https://example.com/two"),
            )
            .is_err()
    );
    assert!(
        editor
            .remove_text_hyperlink(storage_id, hyperlink.id)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
