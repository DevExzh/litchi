use crate::archive::RawMessage;
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::wire::{
    parse_wire_fields, patch_nested_varint_field, patch_varint_field,
    repeated_length_delimited_payloads, transform_length_delimited_field,
};

use super::*;
use crate::text::number_attachment_storage::ATTACHMENT_TABLE_FIELD;
use litchi_iwa_text::number_attachment::{
    TextNumberAttachmentKind, TextNumberAttachmentSettings, TextNumberAttachmentText,
    raw::object_id as native_object_id,
};
use prost::Message;

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

fn storage_message_index(package: &crate::IWorkPackage, storage_id: u64) -> usize {
    locate_attachment_storage(package, storage_id)
        .unwrap()
        .message_index
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
    let location = locate_attachment_storage(&package, storage_id).unwrap();
    let archive_name = location.archive_name.clone();
    let message_index = location.message_index;
    package
        .update_archive(&archive_name, |archive| {
            let storage = archive.object_mut(storage_id).unwrap();
            let original = &storage.messages[message_index];
            let data = transform_length_delimited_field(
                &original.data,
                ATTACHMENT_TABLE_FIELD,
                |table| patch_varint_field(table, 99, false, Some(7)),
            )?;
            storage.replace_message(
                message_index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;

            let object = archive.object_mut(native_object_id(attachment.id)).unwrap();
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
    let message_index = storage_message_index(&package, storage_id);
    let archive = package.archive(&archive_name).unwrap();
    let storage = archive.object(storage_id).unwrap();
    let storage_message = &storage.messages[message_index];
    let table = repeated_length_delimited_payloads(&storage_message.data, ATTACHMENT_TABLE_FIELD)
        .unwrap()[0];
    assert!(
        parse_wire_fields(table)
            .unwrap()
            .iter()
            .any(|field| field.number() == 99)
    );
    let object = archive.object(native_object_id(attachment.id)).unwrap();
    assert!(
        parse_wire_fields(&object.messages[0].data)
            .unwrap()
            .iter()
            .any(|field| field.number() == 88)
    );
    let textual = repeated_length_delimited_payloads(&object.messages[0].data, 1).unwrap()[0];
    assert!(
        parse_wire_fields(textual)
            .unwrap()
            .iter()
            .any(|field| field.number() == 77)
    );
}

#[test]
fn attachment_mutations_use_the_resolved_storage_with_a_2022_style_sibling() {
    let (editor, storage_id) = fixture();
    let mut package = editor.into_package();
    let location = locate_attachment_storage(&package, storage_id).unwrap();
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
    let created = editor
        .insert_text_number_attachment(
            storage_id,
            TextPosition::from_utf16_index(5).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();
    let updated_settings = settings(TextNumberAttachmentKind::PageCount);
    let updated = editor
        .update_text_number_attachment(storage_id, created.id, updated_settings.clone())
        .unwrap();
    assert_eq!(updated.settings, updated_settings);
    editor
        .remove_text_number_attachment(storage_id, created.id)
        .unwrap();

    let package = editor.into_package();
    let location = locate_attachment_storage(&package, storage_id).unwrap();
    let archive = package.archive(&location.archive_name).unwrap();
    let object = archive.object(storage_id).unwrap();
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
fn insertion_relocates_after_text_replacement_moves_existing_attachment() {
    let (mut editor, storage_id) = fixture();
    let first = editor
        .insert_text_number_attachment(
            storage_id,
            TextPosition::from_utf16_index(5).unwrap(),
            settings(TextNumberAttachmentKind::PageNumber),
        )
        .unwrap();

    let second = editor
        .insert_text_number_attachment(
            storage_id,
            TextPosition::from_utf16_index(0).unwrap(),
            settings(TextNumberAttachmentKind::PageCount),
        )
        .unwrap();

    let attachments = editor.text_number_attachments(storage_id).unwrap();
    assert_eq!(second.position, TextPosition::from_utf16_index(0).unwrap());
    assert_eq!(
        attachments
            .iter()
            .find(|attachment| attachment.id == first.id)
            .unwrap()
            .position,
        TextPosition::from_utf16_index(6).unwrap()
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
            let object = archive.object_mut(native_object_id(attachment.id)).unwrap();
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
                .push(native_object_id(attachment.id));
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
