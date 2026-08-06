use crate::archive::RawMessage;
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::highlight_object::validate_plain_highlight_graph;
use crate::text::highlight_storage::locate_storage;
use crate::text::highlight_storage::{HIGHLIGHT_TABLE_FIELD, TABLE_ENTRIES_FIELD};
use crate::wire::{
    append_length_delimited_field, parse_wire_fields, patch_length_delimited_field,
    patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use prost::Message;

use super::*;

const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn fixture() -> (super::super::IWorkTextEditor, u64) {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "Alpha Beta", POSITION, SIZE).unwrap();
    (
        super::super::IWorkTextEditor::from_package(pages.into_package()),
        text_box.storage.object_id,
    )
}

fn range(start: usize, end: usize) -> TextRange {
    TextRange::from_utf16_indexes(start, end).unwrap()
}

fn storage_message_index(package: &crate::IWorkPackage, storage_id: u64) -> usize {
    locate_storage(package, storage_id).unwrap().message_index
}

#[test]
fn unknown_table_entry_and_highlight_fields_survive_updates() {
    let (mut editor, storage_id) = fixture();
    let highlight = editor.add_text_highlight(storage_id, range(0, 5)).unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    let message_index = storage_message_index(&package, storage_id);
    package
        .update_archive(&archive_name, |archive| {
            let storage = archive.object_mut(storage_id).unwrap();
            let original = &storage.messages[message_index];
            let data =
                transform_length_delimited_field(&original.data, HIGHLIGHT_TABLE_FIELD, |table| {
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
            let object = archive.object_mut(highlight.id.object_id()).unwrap();
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
        .update_text_highlight(storage_id, highlight.id, range(6, 10))
        .unwrap();
    let package = editor.into_package();
    let archive = package.archive(&archive_name).unwrap();
    let storage = archive.object(storage_id).unwrap();
    let message = &storage.messages[storage_message_index(&package, storage_id)];
    let table =
        repeated_length_delimited_payloads(&message.data, HIGHLIGHT_TABLE_FIELD).unwrap()[0];
    assert!(
        parse_wire_fields(table)
            .unwrap()
            .iter()
            .any(|field| field.number() == 99)
    );
    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD).unwrap();
    assert!(entries.iter().any(|entry| {
        parse_wire_fields(entry)
            .unwrap()
            .iter()
            .any(|field| field.number() == 77)
    }));
    let object = archive.object(highlight.id.object_id()).unwrap();
    assert!(
        parse_wire_fields(&object.messages[0].data)
            .unwrap()
            .iter()
            .any(|field| field.number() == 88)
    );
}

#[test]
fn duplicate_highlight_tables_fail_without_mutation() {
    let (mut editor, storage_id) = fixture();
    editor.add_text_highlight(storage_id, range(0, 5)).unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    let message_index = storage_message_index(&package, storage_id);
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(storage_id).unwrap();
            let original = &object.messages[message_index];
            let table =
                repeated_length_delimited_payloads(&original.data, HIGHLIGHT_TABLE_FIELD)?[0];
            let mut data = original.data.clone();
            append_length_delimited_field(&mut data, HIGHLIGHT_TABLE_FIELD, table)?;
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
    assert!(editor.text_highlights(storage_id).is_err());
    assert!(editor.add_text_highlight(storage_id, range(6, 10)).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn highlight_updates_use_the_resolved_storage_with_a_2022_style_sibling() {
    let (editor, storage_id) = fixture();
    let mut package = editor.into_package();
    let location = locate_storage(&package, storage_id).unwrap();
    let style_data = tswp::ParagraphStyleArchive {
        super_: crate::protobuf::tss::StyleArchive::default(),
        ..Default::default()
    }
    .encode_to_vec();
    package
        .update_archive(&location.archive_name, |archive| {
            archive
                .object_mut(storage_id)
                .unwrap()
                .push_message(RawMessage {
                    type_: 2_022,
                    data: style_data.clone(),
                })?;
            Ok(())
        })
        .unwrap();

    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let highlight = editor.add_text_highlight(storage_id, range(0, 5)).unwrap();
    editor
        .update_text_highlight(storage_id, highlight.id, range(6, 10))
        .unwrap();
    editor
        .remove_text_highlight(storage_id, highlight.id)
        .unwrap();
    let package = editor.into_package();
    let location = locate_storage(&package, storage_id).unwrap();
    let archive = package.archive(&location.archive_name).unwrap();
    let object = archive.object(storage_id).unwrap();
    assert_eq!(
        object.messages[location.message_index].type_,
        location.message_type
    );
    let sibling = object
        .messages
        .iter()
        .find(|message| message.type_ == 2_022)
        .unwrap();
    assert_eq!(sibling.data, style_data);
}

#[test]
fn highlight_with_an_additional_owner_cannot_be_updated_or_deleted() {
    let (mut editor, storage_id) = fixture();
    let highlight = editor.add_text_highlight(storage_id, range(0, 5)).unwrap();
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
                .push(highlight.id.object_id());
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_highlight(storage_id, highlight.id, range(6, 10))
            .is_err()
    );
    assert!(
        editor
            .remove_text_highlight(storage_id, highlight.id)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn comment_backed_annotations_are_classified_without_mutation() {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "Alpha Beta", POSITION, SIZE).unwrap();
    let storage_id = text_box.storage.object_id;
    let mut package = pages.into_package();
    let created = add_text_highlight(
        &mut package,
        storage_id,
        TextRange::from_utf16_indexes(0, 5).unwrap(),
    )
    .unwrap();
    let location = locate_storage(&package, storage_id).unwrap();
    let archive = package.archive(&location.archive_name).unwrap();
    let object = archive.object(created.id.object_id()).unwrap();
    let graph = validate_plain_highlight_graph(
        &package,
        &location.archive_name,
        created.id.object_id(),
        object,
    )
    .unwrap()
    .unwrap();
    package
        .update_archive(&location.archive_name, |archive| {
            let comment = archive.object_mut(graph.comment_storage_id).unwrap();
            let original = &comment.messages[0];
            let data = patch_length_delimited_field(&original.data, 1, true, Some(b"note"))?;
            comment.replace_message(
                0,
                crate::archive::RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let before = package.to_bytes().unwrap();
    assert!(text_highlights(&package, storage_id).unwrap().is_empty());
    let comments = crate::text::text_comment::text_comments(&package, storage_id).unwrap();
    assert_eq!(comments.len(), 1);
    assert_eq!(
        litchi_iwa_text::comment::raw::comment_id_value(comments[0].id()),
        created.id.object_id()
    );
    assert_eq!(comments[0].body().as_str(), "note");
    assert!(
        remove_text_highlight(&mut package, storage_id, created.id).is_err(),
        "comment-backed annotation must not be discarded"
    );
    assert_eq!(package.to_bytes().unwrap(), before);
}
