use crate::archive::{ArchiveObject, RawMessage};
use crate::pages::PagesEditor;
use crate::wire::{
    parse_wire_fields, patch_nested_varint_field, patch_varint_field,
    repeated_length_delimited_payloads, transform_length_delimited_field,
};

use super::*;
use crate::text::number_attachment_storage::ATTACHMENT_TABLE_FIELD;
use crate::text::number_attachment_types::{
    TextNumberAttachmentKind, TextNumberAttachmentSettings, TextNumberAttachmentText,
};
use crate::text::storage_wire::STORAGE_MESSAGE_TYPES;

fn fixture() -> (super::super::IWorkTextEditor, u64) {
    let pages = PagesEditor::create_with_text("Page ").unwrap();
    let storage_id = pages.body_storage().unwrap().object_id;
    (
        super::super::IWorkTextEditor::from_package(pages.into_package()),
        storage_id,
    )
}

fn settings(kind: TextNumberAttachmentKind) -> TextNumberAttachmentSettings {
    TextNumberAttachmentSettings::new(kind)
        .with_string_equivalent(TextNumberAttachmentText::new("").unwrap())
}

fn storage_message_index(object: &ArchiveObject) -> usize {
    object
        .messages
        .iter()
        .position(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .unwrap()
}

#[test]
fn unknown_table_root_and_nested_fields_survive_updates() {
    let (mut editor, storage_id) = fixture();
    let attachment = editor
        .insert_text_number_attachment(
            storage_id,
            TextPosition::from_utf16_index(5).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_attachment_storage(&package, storage_id)
        .unwrap()
        .archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let storage = archive.object_mut(storage_id).unwrap();
            let index = storage_message_index(storage);
            let original = &storage.messages[index];
            let data = transform_length_delimited_field(
                &original.data,
                ATTACHMENT_TABLE_FIELD,
                |table| patch_varint_field(table, 99, false, Some(7)),
            )?;
            storage.replace_message(
                index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;

            let object = archive.object_mut(attachment.id.object_id()).unwrap();
            let original = &object.messages[0];
            let mut data = patch_varint_field(&original.data, 88, false, Some(9))?;
            data = transform_length_delimited_field(&data, 1, |textual| {
                patch_varint_field(textual, 77, false, Some(42))
            })?;
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
        .update_text_number_attachment(
            storage_id,
            attachment.id,
            settings(TextNumberAttachmentKind::PageCount),
        )
        .unwrap();
    let package = editor.into_package();
    let archive = package.archive(&archive_name).unwrap();
    let storage = archive.object(storage_id).unwrap();
    let storage_message = &storage.messages[storage_message_index(storage)];
    let table = repeated_length_delimited_payloads(&storage_message.data, ATTACHMENT_TABLE_FIELD)
        .unwrap()[0];
    assert!(
        parse_wire_fields(table)
            .unwrap()
            .iter()
            .any(|field| field.number == 99)
    );
    let object = archive.object(attachment.id.object_id()).unwrap();
    assert!(
        parse_wire_fields(&object.messages[0].data)
            .unwrap()
            .iter()
            .any(|field| field.number == 88)
    );
    let textual = repeated_length_delimited_payloads(&object.messages[0].data, 1).unwrap()[0];
    assert!(
        parse_wire_fields(textual)
            .unwrap()
            .iter()
            .any(|field| field.number == 77)
    );
}

#[test]
fn missing_kind_and_additional_owner_fail_transactionally() {
    let (mut editor, storage_id) = fixture();
    let attachment = editor
        .insert_text_number_attachment(
            storage_id,
            TextPosition::from_utf16_index(5).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_attachment_storage(&package, storage_id)
        .unwrap()
        .archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(attachment.id.object_id()).unwrap();
            let original = &object.messages[0];
            let data = patch_nested_varint_field(&original.data, &[1, 2], true, None)?;
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
    let editor = super::super::IWorkTextEditor::from_package(package);
    assert!(editor.text_number_attachments(storage_id).is_err());

    let (mut editor, storage_id) = fixture();
    let attachment = editor
        .insert_text_number_attachment(
            storage_id,
            TextPosition::from_utf16_index(5).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_attachment_storage(&package, storage_id)
        .unwrap()
        .archive_name;
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
                .push(attachment.id.object_id());
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_number_attachment(
                storage_id,
                attachment.id,
                settings(TextNumberAttachmentKind::PageCount),
            )
            .is_err()
    );
    assert!(
        editor
            .remove_text_number_attachment(storage_id, attachment.id)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
