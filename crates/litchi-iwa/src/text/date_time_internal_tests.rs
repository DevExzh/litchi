use crate::archive::RawMessage;
use crate::pages::PagesEditor;
use crate::protobuf::tswp;
use crate::wire::{
    parse_wire_fields, patch_fixed64_field, patch_varint_field, repeated_length_delimited_payloads,
    transform_length_delimited_field,
};

use super::*;
use crate::text::hyperlink_storage::SMART_FIELD_TABLE_FIELD;
use litchi_iwa_text::date_time::{Format, FormatterStyle, Instant, LocaleIdentifier, Settings};
use prost::Message;

const DATE_FIELD: u32 = 8;

fn settings(seconds: f64) -> Settings {
    Settings::fixed(
        Format::new("EEEE, MMMM d, y").unwrap(),
        LocaleIdentifier::new("en_US").unwrap(),
        Instant::from_reference_date_seconds(seconds).unwrap(),
    )
    .with_styles(FormatterStyle::Full, FormatterStyle::None)
    .unwrap()
}

fn fixture() -> (super::super::IWorkTextEditor, u64) {
    let pages = PagesEditor::create_with_text("Friday, July 17, 2026").unwrap();
    let storage_id = pages.body_storage().unwrap().object_id;
    (
        super::super::IWorkTextEditor::from_package(pages.into_package()),
        storage_id,
    )
}

fn whole_range() -> TextRange {
    TextRange::from_utf16_indexes(0, "Friday, July 17, 2026".encode_utf16().count()).unwrap()
}

fn storage_message_index(package: &crate::IWorkPackage, storage_id: u64) -> usize {
    locate_storage(package, storage_id, SMART_FIELD_TABLE)
        .unwrap()
        .message_index
}

#[test]
fn unknown_table_payload_and_nested_date_fields_survive_updates() {
    let (mut editor, storage_id) = fixture();
    let field = editor
        .add_text_date_time_field(storage_id, whole_range(), settings(805_965_335.0))
        .unwrap();
    let mut package = editor.into_package();
    let location = locate_storage(&package, storage_id, SMART_FIELD_TABLE).unwrap();
    let archive_name = location.archive_name.clone();
    let message_index = location.message_index;
    package
        .update_archive(&archive_name, |archive| {
            let storage = archive.object_mut(storage_id).unwrap();
            let original = &storage.messages[message_index];
            let data = transform_length_delimited_field(
                &original.data,
                SMART_FIELD_TABLE_FIELD,
                |table| patch_varint_field(table, 99, false, Some(7)),
            )?;
            storage.replace_message(
                message_index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            let object = archive.object_mut(field.id.object_id()).unwrap();
            let original = &object.messages[0];
            let mut data = patch_varint_field(&original.data, 88, false, Some(9))?;
            data = transform_length_delimited_field(&data, DATE_FIELD, |date| {
                patch_varint_field(date, 77, false, Some(42))
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
        .update_text_date_time_field(storage_id, field.id, whole_range(), settings(900_000_000.5))
        .unwrap();
    let package = editor.into_package();
    let message_index = storage_message_index(&package, storage_id);
    let archive = package.archive(&archive_name).unwrap();
    let storage = archive.object(storage_id).unwrap();
    let message = &storage.messages[message_index];
    let table =
        repeated_length_delimited_payloads(&message.data, SMART_FIELD_TABLE_FIELD).unwrap()[0];
    assert!(
        parse_wire_fields(table)
            .unwrap()
            .iter()
            .any(|wire| wire.number() == 99)
    );
    let object = archive.object(field.id.object_id()).unwrap();
    assert!(
        parse_wire_fields(&object.messages[0].data)
            .unwrap()
            .iter()
            .any(|wire| wire.number() == 88)
    );
    let date = repeated_length_delimited_payloads(&object.messages[0].data, DATE_FIELD).unwrap()[0];
    assert!(
        parse_wire_fields(date)
            .unwrap()
            .iter()
            .any(|wire| wire.number() == 77)
    );
}

#[test]
fn date_time_mutations_use_the_resolved_storage_with_a_2022_style_sibling() {
    let (editor, storage_id) = fixture();
    let mut package = editor.into_package();
    let location = locate_storage(&package, storage_id, SMART_FIELD_TABLE).unwrap();
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
    let field = editor
        .add_text_date_time_field(storage_id, whole_range(), settings(805_965_335.0))
        .unwrap();
    editor
        .update_text_date_time_field(storage_id, field.id, whole_range(), settings(900_000_000.5))
        .unwrap();
    editor
        .remove_text_date_time_field(storage_id, field.id)
        .unwrap();

    let package = editor.into_package();
    let location = locate_storage(&package, storage_id, SMART_FIELD_TABLE).unwrap();
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
fn malformed_instant_and_additional_owner_fail_without_mutation() {
    let (mut editor, storage_id) = fixture();
    let field = editor
        .add_text_date_time_field(storage_id, whole_range(), settings(805_965_335.0))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id, SMART_FIELD_TABLE)
        .unwrap()
        .archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(field.id.object_id()).unwrap();
            let original = &object.messages[0];
            let data = transform_length_delimited_field(&original.data, DATE_FIELD, |date| {
                patch_fixed64_field(date, 1, true, Some(f64::NAN.to_bits()))
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
    let editor = super::super::IWorkTextEditor::from_package(package);
    assert!(editor.text_date_time_fields(storage_id).is_err());

    let (mut editor, storage_id) = fixture();
    let field = editor
        .add_text_date_time_field(storage_id, whole_range(), settings(805_965_335.0))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id, SMART_FIELD_TABLE)
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
                .push(field.id.object_id());
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_date_time_field(
                storage_id,
                field.id,
                whole_range(),
                settings(900_000_000.0),
            )
            .is_err()
    );
    assert!(
        editor
            .remove_text_date_time_field(storage_id, field.id)
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
