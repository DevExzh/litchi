//! Native paragraph-alignment inheritance and copy-on-write style updates.

use std::collections::HashSet;

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, release_package_identifier_suffix,
    remove_component_external_references_to_object, remove_component_object_uuids,
    set_package_last_object_identifier,
};
use crate::protobuf::{tsp, tss, tswp};
use crate::shapes::{insert_style_variation, remove_style_variation};
use crate::wire::{
    parse_wire_fields, patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::style::TextAlignment;

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

const STORAGE_PARAGRAPH_STYLE_TABLE_FIELD: u32 = 5;
const ATTRIBUTE_TABLE_ENTRIES_FIELD: u32 = 1;
const ATTRIBUTE_CHARACTER_INDEX_FIELD: u32 = 1;
const ATTRIBUTE_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_CHARACTER_PROPERTIES_FIELD: u32 = 11;
const STYLE_PARAGRAPH_PROPERTIES_FIELD: u32 = 12;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const PARAGRAPH_ALIGNMENT_FIELD: u32 = 1;

pub(super) fn paragraph_alignment(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<TextAlignment> {
    let (_, _, storage) = storage_payload(package, storage_id)?;
    let style_id = uniform_paragraph_style_id(&storage, storage_id)?;
    inherited_paragraph_alignment(package, style_id)
}

pub(super) fn set_paragraph_alignment(
    package: &mut IWorkPackage,
    storage_id: u64,
    alignment: TextAlignment,
) -> Result<()> {
    if paragraph_alignment(package, storage_id)? == alignment {
        return Ok(());
    }

    let (storage_archive_name, _, storage) = storage_payload(package, storage_id)?;
    let old_style_id = uniform_paragraph_style_id(&storage, storage_id)?;
    let (style_archive_name, old_style_message, old_style) =
        paragraph_style_payload(package, old_style_id)?;
    let stylesheet_id = old_style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {old_style_id} has no stylesheet"
            ))
        })?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if stylesheet_archive_name != style_archive_name {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {old_style_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }

    let direct = direct_alignment_variation(&old_style, &old_style_message.data)?;
    if direct && paragraph_style_is_exclusive(package, old_style_id)? {
        let parent_style_id = required_parent_style_id(&old_style, old_style_id)?;
        let replacement = paragraph_style_variation_object(
            old_style_id,
            parent_style_id,
            stylesheet_id,
            alignment,
        )?;
        replace_paragraph_style_variation(package, &style_archive_name, old_style_id, replacement)?;
        return Ok(());
    }

    let new_style_id = next_object_identifier(package)?;
    let new_style =
        paragraph_style_variation_object(new_style_id, old_style_id, stylesheet_id, alignment)?;
    let mut staged = package.clone();
    patch_storage_paragraph_style_reference(
        &mut staged,
        &storage_archive_name,
        storage_id,
        old_style_id,
        new_style_id,
    )?;
    insert_style_variation(
        &mut staged,
        &style_archive_name,
        stylesheet_id,
        old_style_id,
        new_style_id,
        new_style,
    )?;
    if let Some(style_component) = component_identifier_for_entry(&staged, &style_archive_name)? {
        add_component_object_uuids(&mut staged, style_component, &[new_style_id])?;
        if let Some(storage_component) =
            component_identifier_for_entry(&staged, &storage_archive_name)?
            && storage_component != style_component
        {
            add_component_external_reference(
                &mut staged,
                storage_component,
                style_component,
                new_style_id,
            )?;
        }
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    if paragraph_alignment(&staged, storage_id)? != alignment {
        return Err(Error::InvalidFormat(
            "iWork paragraph-alignment update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_alignment(
    package: &mut IWorkPackage,
    storage_id: u64,
) -> Result<bool> {
    let (storage_archive_name, _, storage) = storage_payload(package, storage_id)?;
    let style_id = uniform_paragraph_style_id(&storage, storage_id)?;
    let (style_archive_name, style_message, style) = paragraph_style_payload(package, style_id)?;
    if !direct_alignment_variation(&style, &style_message.data)?
        || !paragraph_style_is_exclusive(package, style_id)?
    {
        return Ok(false);
    }
    let parent_style_id = required_parent_style_id(&style, style_id)?;
    let stylesheet_id = style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph-style variation {style_id} has no stylesheet"
            ))
        })?;

    let mut staged = package.clone();
    patch_storage_paragraph_style_reference(
        &mut staged,
        &storage_archive_name,
        storage_id,
        style_id,
        parent_style_id,
    )?;
    remove_style_variation(
        &mut staged,
        &style_archive_name,
        stylesheet_id,
        parent_style_id,
        style_id,
    )?;
    if let Some(style_component) = component_identifier_for_entry(&staged, &style_archive_name)? {
        remove_component_object_uuids(&mut staged, style_component, &[style_id])?;
        remove_component_external_references_to_object(&mut staged, style_component, style_id)?;
        if let Some(storage_component) =
            component_identifier_for_entry(&staged, &storage_archive_name)?
            && storage_component != style_component
        {
            add_component_external_reference(
                &mut staged,
                storage_component,
                style_component,
                parent_style_id,
            )?;
        }
    }
    release_package_identifier_suffix(&mut staged, &[style_id])?;
    let expected = inherited_paragraph_alignment(&staged, parent_style_id)?;
    if paragraph_alignment(&staged, storage_id)? != expected {
        return Err(Error::InvalidFormat(
            "iWork paragraph-alignment reset failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(true)
}

fn inherited_paragraph_alignment(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextAlignment> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(TextAlignment::Natural);
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style inheritance cycles at {identifier}"
            )));
        }
        let (_, _, style) = paragraph_style_payload(package, identifier)?;
        if let Some(value) = style
            .para_properties
            .as_ref()
            .and_then(|properties| properties.alignment)
        {
            return TextAlignment::from_native_value(value);
        }
        style_id = style.super_.parent.map(|parent| parent.identifier);
    }
    Err(Error::InvalidFormat(format!(
        "iWork paragraph style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}

fn storage_payload(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<(String, RawMessage, tswp::StorageArchive)> {
    let archive_name = object_archive_name(package, storage_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
    })?;
    let payloads = object
        .messages
        .iter()
        .filter_map(|message| {
            STORAGE_MESSAGE_TYPES
                .contains(&message.type_)
                .then(|| {
                    tswp::StorageArchive::decode(message.data.as_slice())
                        .ok()
                        .map(|storage| (message.clone(), storage))
                })
                .flatten()
        })
        .collect::<Vec<_>>();
    let [(message, storage)] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must have exactly one writable payload"
        )));
    };
    let table_count = repeated_length_delimited_payloads(
        message.data.as_slice(),
        STORAGE_PARAGRAPH_STYLE_TABLE_FIELD,
    )?
    .len();
    if table_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one paragraph-style table, found {table_count}"
        )));
    }
    Ok((archive_name, message.clone(), storage.clone()))
}

fn uniform_paragraph_style_id(storage: &tswp::StorageArchive, storage_id: u64) -> Result<u64> {
    let entries = storage
        .table_para_style
        .as_ref()
        .map(|table| table.entries.as_slice())
        .unwrap_or_default();
    let [entry] = entries else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must have one uniform paragraph-style boundary"
        )));
    };
    if entry.character_index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph style must begin at UTF-16 index zero"
        )));
    }
    entry
        .object
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no uniform paragraph style"
            ))
        })
}

fn paragraph_style_payload(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<(String, RawMessage, tswp::ParagraphStyleArchive)> {
    let archive_name = object_archive_name(package, style_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(style_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork paragraph style {style_id} is missing"))
    })?;
    let payloads = object
        .messages
        .iter()
        .filter(|message| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
        .filter_map(|message| {
            tswp::ParagraphStyleArchive::decode(message.data.as_slice())
                .ok()
                .map(|style| (message.clone(), style))
        })
        .collect::<Vec<_>>();
    let [(message, style)] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {style_id} must have exactly one paragraph-style payload"
        )));
    };
    Ok((archive_name, message.clone(), style.clone()))
}

fn direct_alignment_variation(style: &tswp::ParagraphStyleArchive, raw: &[u8]) -> Result<bool> {
    let Some(properties) = style.para_properties.as_ref() else {
        return Ok(false);
    };
    let semantic = properties.alignment.is_some()
        && style.override_count == Some(1)
        && style
            .char_properties
            .as_ref()
            .is_some_and(|value| *value == tswp::CharacterStylePropertiesArchive::default())
        && style.super_.name.is_none()
        && style.super_.style_identifier.is_none()
        && style.super_.parent.is_some()
        && style.super_.is_variation == Some(true)
        && style.super_.stylesheet.is_some();
    if !semantic {
        return Ok(false);
    }
    let super_raw = required_payload(raw, STYLE_SUPER_FIELD, "paragraph style")?;
    let character_raw = required_payload(
        raw,
        STYLE_CHARACTER_PROPERTIES_FIELD,
        "paragraph character properties",
    )?;
    let paragraph_raw = required_payload(
        raw,
        STYLE_PARAGRAPH_PROPERTIES_FIELD,
        "paragraph properties",
    )?;
    Ok(has_exact_fields(
        raw,
        &[
            STYLE_SUPER_FIELD,
            STYLE_OVERRIDE_COUNT_FIELD,
            STYLE_CHARACTER_PROPERTIES_FIELD,
            STYLE_PARAGRAPH_PROPERTIES_FIELD,
        ],
    )? && has_exact_fields(
        super_raw,
        &[
            STYLE_PARENT_FIELD,
            STYLE_VARIATION_FIELD,
            STYLE_STYLESHEET_FIELD,
        ],
    )? && has_exact_fields(character_raw, &[])?
        && has_exact_fields(paragraph_raw, &[PARAGRAPH_ALIGNMENT_FIELD])?)
}

fn paragraph_style_is_exclusive(package: &IWorkPackage, style_id: u64) -> Result<bool> {
    let mut storage_references = 0usize;
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            for message in &object.messages {
                if STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && let Ok(storage) = tswp::StorageArchive::decode(message.data.as_slice())
                {
                    storage_references += storage
                        .table_para_style
                        .iter()
                        .flat_map(|table| &table.entries)
                        .filter(|entry| {
                            entry
                                .object
                                .as_ref()
                                .is_some_and(|reference| reference.identifier == style_id)
                        })
                        .count();
                }
                if message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                    && let Ok(style) = tswp::ParagraphStyleArchive::decode(message.data.as_slice())
                    && style
                        .super_
                        .parent
                        .as_ref()
                        .is_some_and(|parent| parent.identifier == style_id)
                {
                    return Ok(false);
                }
            }
        }
    }
    Ok(storage_references == 1)
}

fn paragraph_style_variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    alignment: TextAlignment,
) -> Result<ArchiveObject> {
    let data = tswp::ParagraphStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        char_properties: Some(tswp::CharacterStylePropertiesArchive::default()),
        para_properties: Some(tswp::ParagraphStylePropertiesArchive {
            alignment: Some(alignment.native_value()),
            ..Default::default()
        }),
    }
    .encode_to_vec();
    tswp::ParagraphStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    object.archive_info.message_infos[0]
        .object_references
        .push(parent_style_id);
    Ok(object)
}

fn replace_paragraph_style_variation(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_id: u64,
    mut replacement: ArchiveObject,
) -> Result<()> {
    let message = replacement.messages.pop().ok_or_else(|| {
        Error::InvalidFormat("replacement paragraph style has no payload".to_owned())
    })?;
    if !replacement.messages.is_empty() {
        return Err(Error::InvalidFormat(
            "replacement paragraph style has multiple payloads".to_owned(),
        ));
    }
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork paragraph style {style_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, candidate)| candidate.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style {style_id} must have exactly one payload"
            )));
        };
        object.replace_message(*index, message)?;
        Ok(())
    })
}

fn patch_storage_paragraph_style_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| {
                STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && tswp::StorageArchive::decode(message.data.as_slice()).is_ok()
            })
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        };
        let original = &object.messages[*index];
        let data = transform_length_delimited_field(
            &original.data,
            STORAGE_PARAGRAPH_STYLE_TABLE_FIELD,
            |table| {
                let entries = repeated_length_delimited_payloads(
                    table,
                    ATTRIBUTE_TABLE_ENTRIES_FIELD,
                )?;
                let [entry] = entries.as_slice() else {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} must have one uniform paragraph-style boundary"
                    )));
                };
                let character_index = required_varint(
                    entry,
                    ATTRIBUTE_CHARACTER_INDEX_FIELD,
                    "paragraph-style character index",
                )?;
                if character_index != 0 {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} paragraph style must begin at index zero"
                    )));
                }
                let object_data = required_payload(
                    entry,
                    ATTRIBUTE_OBJECT_FIELD,
                    "paragraph-style reference",
                )?;
                let identifier = required_varint(
                    object_data,
                    REFERENCE_IDENTIFIER_FIELD,
                    "paragraph-style identifier",
                )?;
                if identifier != old_style_id {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} paragraph style changed unexpectedly"
                    )));
                }
                let patched = transform_length_delimited_field(
                    entry,
                    ATTRIBUTE_OBJECT_FIELD,
                    |reference| {
                        patch_varint_field(
                            reference,
                            REFERENCE_IDENTIFIER_FIELD,
                            true,
                            Some(new_style_id),
                        )
                    },
                )?;
                rewrite_repeated_length_delimited_fields(
                    table,
                    ATTRIBUTE_TABLE_ENTRIES_FIELD,
                    &[patched],
                )
            },
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        let mut replaced = 0usize;
        for reference in &mut info.object_references {
            if *reference == old_style_id {
                *reference = new_style_id;
                replaced += 1;
            }
        }
        for field in &mut info.field_infos {
            for reference in &mut field.object_references {
                if *reference == old_style_id {
                    *reference = new_style_id;
                }
            }
        }
        if replaced != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} metadata contains {replaced} paragraph-style references"
            )));
        }
        Ok(())
    })
}

fn required_parent_style_id(style: &tswp::ParagraphStyleArchive, style_id: u64) -> Result<u64> {
    style
        .super_
        .parent
        .as_ref()
        .map(|parent| parent.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph-style variation {style_id} has no parent"
            ))
        })
}

fn object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_some()
            && found.replace(name.to_owned()).is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork object {identifier} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork object {identifier} is missing")))
}

fn has_exact_fields(data: &[u8], expected: &[u32]) -> Result<bool> {
    let mut actual = parse_wire_fields(data)?
        .into_iter()
        .map(|field| field.number)
        .collect::<Vec<_>>();
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    Ok(actual == expected)
}

fn required_payload<'a>(data: &'a [u8], field: u32, context: &str) -> Result<&'a [u8]> {
    let payloads = repeated_length_delimited_payloads(data, field)?;
    let [payload] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain field {field} exactly once"
        )));
    };
    Ok(payload)
}

fn required_varint(data: &[u8], field_number: u32, context: &str) -> Result<u64> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number && field.wire_type == 0)
        .collect::<Vec<_>>();
    let [field] = matches.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain varint field {field_number} exactly once"
        )));
    };
    let (value, length) = crate::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
        .map_err(|error| Error::InvalidFormat(format!("invalid {context}: {error}")))?;
    if field.key_end + length != field.end {
        return Err(Error::InvalidFormat(format!(
            "{context} has trailing varint bytes"
        )));
    }
    Ok(value)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::pages::PagesEditor;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use crate::text::{TextColumnCount, TextColumns};

    #[test]
    fn all_native_alignment_values_are_strict_and_reversible() {
        for alignment in [
            TextAlignment::Natural,
            TextAlignment::Right,
            TextAlignment::Center,
            TextAlignment::Justified,
            TextAlignment::Left,
        ] {
            assert_eq!(
                TextAlignment::from_native_value(alignment.native_value()).unwrap(),
                alignment
            );
        }
        assert!(TextAlignment::from_native_value(-1).is_err());
        assert!(TextAlignment::from_native_value(5).is_err());
    }

    #[test]
    fn pages_text_box_alignment_is_independent_replaceable_and_resettable() {
        let mut editor = PagesEditor::create_with_text("Alignment").unwrap();
        let first = editor
            .add_text_box(
                9,
                "First paragraph",
                DrawablePoint { x: 20.0, y: 40.0 },
                DrawableSize {
                    width: 240.0,
                    height: 100.0,
                },
            )
            .unwrap();
        let second = editor
            .add_text_box(
                10,
                "Second paragraph",
                DrawablePoint { x: 40.0, y: 160.0 },
                DrawableSize {
                    width: 240.0,
                    height: 100.0,
                },
            )
            .unwrap();
        let columns = TextColumns::equal(TextColumnCount::new(2).unwrap(), None);
        editor
            .set_text_box_columns(first.drawable_object_id, &columns)
            .unwrap();

        assert_eq!(
            editor
                .text_box_paragraph_alignment(first.drawable_object_id)
                .unwrap(),
            TextAlignment::Natural
        );
        editor
            .set_text_box_paragraph_alignment(first.drawable_object_id, TextAlignment::Center)
            .unwrap();
        assert_eq!(
            editor
                .text_box_paragraph_alignment(second.drawable_object_id)
                .unwrap(),
            TextAlignment::Natural
        );
        editor
            .set_text_box_paragraph_alignment(first.drawable_object_id, TextAlignment::Left)
            .unwrap();

        let mut reopened = PagesEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .text_box_paragraph_alignment(first.drawable_object_id)
                .unwrap(),
            TextAlignment::Left
        );
        assert_eq!(
            reopened.text_box_columns(first.drawable_object_id).unwrap(),
            columns
        );
        assert!(
            reopened
                .reset_text_box_paragraph_alignment(first.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .text_box_paragraph_alignment(first.drawable_object_id)
                .unwrap(),
            TextAlignment::Natural
        );
        assert!(
            !reopened
                .reset_text_box_paragraph_alignment(first.drawable_object_id)
                .unwrap()
        );
    }

    #[test]
    fn numbers_text_box_alignment_round_trips_and_resets() {
        let mut editor = NumbersDocumentBuilder::new().build().unwrap();
        let sheet_id = editor.sheets().unwrap()[0].object_id;
        let created = editor
            .add_sheet_text_box(
                sheet_id,
                "Right aligned",
                DrawablePoint { x: 20.0, y: 200.0 },
                DrawableSize {
                    width: 240.0,
                    height: 100.0,
                },
            )
            .unwrap();
        editor
            .set_sheet_text_box_paragraph_alignment(
                sheet_id,
                created.drawable_object_id,
                TextAlignment::Right,
            )
            .unwrap();
        let mut reopened =
            crate::numbers::NumbersEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .sheet_text_box_paragraph_alignment(sheet_id, created.drawable_object_id)
                .unwrap(),
            TextAlignment::Right
        );
        assert!(
            reopened
                .reset_sheet_text_box_paragraph_alignment(sheet_id, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .sheet_text_box_paragraph_alignment(sheet_id, created.drawable_object_id)
                .unwrap(),
            TextAlignment::Natural
        );
    }

    #[test]
    fn keynote_text_box_alignment_round_trips_and_resets() {
        let mut editor = KeynoteDocumentBuilder::new().build().unwrap();
        let created = editor
            .add_slide_text_box(
                0,
                "Justified paragraph",
                DrawablePoint { x: 80.0, y: 500.0 },
                DrawableSize {
                    width: 500.0,
                    height: 100.0,
                },
            )
            .unwrap();
        editor
            .set_slide_text_box_paragraph_alignment(
                0,
                created.drawable_object_id,
                TextAlignment::Justified,
            )
            .unwrap();
        let mut reopened =
            crate::keynote::KeynoteEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reopened
                .slide_text_box_paragraph_alignment(0, created.drawable_object_id)
                .unwrap(),
            TextAlignment::Justified
        );
        assert!(
            reopened
                .reset_slide_text_box_paragraph_alignment(0, created.drawable_object_id)
                .unwrap()
        );
        assert_eq!(
            reopened
                .slide_text_box_paragraph_alignment(0, created.drawable_object_id)
                .unwrap(),
            TextAlignment::Natural
        );
    }

    #[test]
    fn multiple_paragraph_boundaries_are_rejected_transactionally() {
        let mut pages = PagesEditor::create_with_text("Alignment").unwrap();
        let created = pages
            .add_text_box(
                9,
                "Two paragraphs",
                DrawablePoint { x: 20.0, y: 40.0 },
                DrawableSize {
                    width: 240.0,
                    height: 100.0,
                },
            )
            .unwrap();
        let storage_id = created.storage.object_id;
        let mut package = pages.into_package();
        let (archive_name, _, _) = storage_payload(&package, storage_id).unwrap();
        package
            .update_archive(&archive_name, |archive| {
                let object = archive.object_mut(storage_id).unwrap();
                let index = object
                    .messages
                    .iter()
                    .position(|message| {
                        STORAGE_MESSAGE_TYPES.contains(&message.type_)
                            && tswp::StorageArchive::decode(message.data.as_slice()).is_ok()
                    })
                    .unwrap();
                let message_type = object.messages[index].type_;
                let mut storage =
                    tswp::StorageArchive::decode(object.messages[index].data.as_slice()).unwrap();
                let mut second = storage.table_para_style.as_ref().unwrap().entries[0];
                second.character_index = 1;
                storage
                    .table_para_style
                    .as_mut()
                    .unwrap()
                    .entries
                    .push(second);
                object
                    .replace_message(
                        index,
                        RawMessage {
                            type_: message_type,
                            data: storage.encode_to_vec(),
                        },
                    )
                    .unwrap();
                Ok(())
            })
            .unwrap();
        let mut editor = super::super::editor::IWorkTextEditor::from_package(package);
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_paragraph_alignment(storage_id, TextAlignment::Center)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
