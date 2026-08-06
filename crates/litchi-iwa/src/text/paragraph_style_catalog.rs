//! Theme-visible named paragraph-style creation.

use prost::Message;

use crate::archive::{Archive, RawMessage};
use crate::package_metadata::{
    add_component_external_reference, add_component_object_uuids, component_identifier_for_entry,
    next_object_identifier, set_package_last_object_identifier,
};
use crate::protobuf::{tsp, tss};
use crate::text::storage_wire::update_parsed_archive;
use crate::wire::{
    append_repeated_length_delimited_field, patch_nested_length_delimited_field,
    patch_nested_varint_field,
};
use crate::{Error, IWorkPackage, IWorkThemeArchive, Result};

use super::paragraph_alignment::native;
use super::style_registry::object_archive_name;
use litchi_iwa_text::paragraph::style::{
    NamedParagraphStyle, ParagraphStyleId, ParagraphStyleName,
    raw::{from_native_id, native_id},
};

const THEME_MESSAGE_TYPES: &[u32] = &[10, 10_001, 12_009];
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const STYLESHEET_IDENTIFIER_MAP_FIELD: u32 = 2;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_NAME_FIELD: u32 = 1;
const STYLE_IDENTIFIER_FIELD: u32 = 2;
const STYLE_VARIATION_FIELD: u32 = 4;
const GENERATED_STYLE_IDENTIFIER_PREFIX: &str = "com.litchi.paragraph-style.";

#[derive(Debug)]
pub(super) struct ThemeLocation {
    pub(super) archive_name: String,
    pub(super) object_id: u64,
    pub(super) message_type: u32,
}

pub(super) fn create_named_paragraph_style(
    package: &mut IWorkPackage,
    first_style_id: u64,
    source: ParagraphStyleId,
    name: ParagraphStyleName,
) -> Result<NamedParagraphStyle> {
    native::validate_named_paragraph_style(package, first_style_id, source)?;
    let existing = native::named_paragraph_styles(package, first_style_id)?;
    if existing
        .iter()
        .any(|style| style.name().as_str() == name.as_str())
    {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style name {:?} already exists",
            name.as_str()
        )));
    }

    let source_id = native_id(source);
    let native::LocatedParagraphStyle {
        location: source_location,
        archive: source_archive,
        package_revision,
    } = native::locate_style_with_archive(package, source_id)?;
    let stylesheet_id = native::stylesheet_id(&source_location.style, source_id)?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if source_location.archive_name != stylesheet_archive_name {
        return Err(Error::InvalidFormat(format!(
            "named iWork paragraph style {source_id} is not stored with stylesheet {stylesheet_id}"
        )));
    }
    let theme_locations = locate_themes(package, stylesheet_id, source_id)?;
    let new_style_id = next_object_identifier(package)?;
    let style_identifier = format!("{GENERATED_STYLE_IDENTIFIER_PREFIX}{new_style_id}");
    let new_style = clone_named_style(
        &source_archive,
        &source_location,
        new_style_id,
        name.as_str(),
        &style_identifier,
    )?;

    let mut staged = package.clone();
    insert_named_style_with_archive(
        &mut staged,
        &stylesheet_archive_name,
        stylesheet_id,
        new_style_id,
        &style_identifier,
        new_style,
        source_archive,
        package_revision,
    )?;
    for location in &theme_locations {
        append_theme_preset(&mut staged, location, stylesheet_id, new_style_id)?;
    }
    if let Some(style_component) =
        component_identifier_for_entry(&staged, &stylesheet_archive_name)?
    {
        add_component_object_uuids(&mut staged, style_component, &[new_style_id])?;
        for location in &theme_locations {
            if let Some(theme_component) =
                component_identifier_for_entry(&staged, &location.archive_name)?
                && theme_component != style_component
            {
                add_component_external_reference(
                    &mut staged,
                    theme_component,
                    style_component,
                    new_style_id,
                )?;
            }
        }
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;

    let created =
        NamedParagraphStyle::from_owned(from_native_id(new_style_id)?, name.as_str().to_owned())?;
    let matches = native::named_paragraph_styles(&staged, first_style_id)?
        .into_iter()
        .filter(|style| style == &created)
        .count();
    if matches != 1 {
        return Err(Error::InvalidFormat(
            "named iWork paragraph style creation failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(created)
}

pub(super) fn locate_themes(
    package: &IWorkPackage,
    stylesheet_id: u64,
    source_style_id: u64,
) -> Result<Vec<ThemeLocation>> {
    let mut locations = Vec::new();
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        for object in &archive.objects {
            let Some(object_id) = object.archive_info.identifier else {
                continue;
            };
            for message in &object.messages {
                if !THEME_MESSAGE_TYPES.contains(&message.type_) {
                    continue;
                }
                let theme = IWorkThemeArchive::decode(&message.data)?;
                if theme
                    .base
                    .document_stylesheet
                    .as_ref()
                    .map(|reference| reference.identifier)
                    != Some(stylesheet_id)
                {
                    continue;
                }
                let contains_source = theme.extensions.text.as_ref().is_some_and(|text| {
                    text.paragraph_style_presets
                        .iter()
                        .any(|reference| reference.identifier == source_style_id)
                });
                if contains_source {
                    locations.push(ThemeLocation {
                        archive_name: archive_name.to_owned(),
                        object_id,
                        message_type: message.type_,
                    });
                }
            }
        }
    }
    if locations.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet {stylesheet_id} has no theme containing paragraph style {source_style_id}"
        )));
    }
    Ok(locations)
}

fn clone_named_style(
    archive: &Archive,
    source_location: &native::ParagraphStyleLocation,
    new_id: u64,
    name: &str,
    style_identifier: &str,
) -> Result<crate::archive::ArchiveObject> {
    let source_id = source_location.object_id;
    let source = archive.object(source_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork paragraph style {source_id} is missing"))
    })?;
    let indexes = source
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| {
            (message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE).then_some(index)
        })
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {source_id} must have exactly one paragraph-style payload"
        )));
    };
    let message_index = *message_index;
    let decoded = &source_location.style;
    let mut data = patch_nested_length_delimited_field(
        &source.messages[message_index].data,
        &[STYLE_SUPER_FIELD, STYLE_NAME_FIELD],
        decoded.super_.name.is_some(),
        Some(name.as_bytes()),
    )?;
    data = patch_nested_length_delimited_field(
        &data,
        &[STYLE_SUPER_FIELD, STYLE_IDENTIFIER_FIELD],
        decoded.super_.style_identifier.is_some(),
        Some(style_identifier.as_bytes()),
    )?;
    data = patch_nested_varint_field(
        &data,
        &[STYLE_SUPER_FIELD, STYLE_VARIATION_FIELD],
        decoded.super_.is_variation.is_some(),
        None,
    )?;

    let mut cloned = crate::archive::ArchiveObject::new(new_id, source.messages.clone())?;
    cloned.archive_info.message_infos = source.archive_info.message_infos.clone();
    cloned.archive_info.should_merge = source.archive_info.should_merge;
    cloned.replace_message(
        message_index,
        RawMessage {
            type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
            data,
        },
    )?;
    Ok(cloned)
}

fn insert_named_style_with_archive(
    package: &mut IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    style_id: u64,
    style_identifier: &str,
    style: crate::archive::ArchiveObject,
    archive: Archive,
    package_revision: u64,
) -> Result<()> {
    if package.mutation_revision() != package_revision {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {style_id} package changed unexpectedly"
        )));
    }
    update_parsed_archive(package, archive_name, archive, |archive| {
        insert_named_style_in_archive(archive, stylesheet_id, style_id, style_identifier, style)
    })
}

fn insert_named_style_in_archive(
    archive: &mut Archive,
    stylesheet_id: u64,
    style_id: u64,
    style_identifier: &str,
    style: crate::archive::ArchiveObject,
) -> Result<()> {
    let stylesheet = archive.object_mut(stylesheet_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork stylesheet {stylesheet_id} is missing"))
    })?;
    let indexes = stylesheet
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| (message.type_ == STYLESHEET_MESSAGE_TYPE).then_some(index))
        .collect::<Vec<_>>();
    let [message_index] = indexes.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet {stylesheet_id} must have exactly one stylesheet payload"
        )));
    };
    let message_index = *message_index;
    let decoded =
        tss::StylesheetArchive::decode(stylesheet.messages[message_index].data.as_slice())?;
    if decoded
        .styles
        .iter()
        .any(|value| value.identifier == style_id)
        || decoded
            .identifier_to_style_map
            .iter()
            .any(|entry| entry.identifier == style_identifier)
    {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet already contains paragraph style {style_id}"
        )));
    }
    let reference = tsp::Reference {
        identifier: style_id,
        ..Default::default()
    };
    let entry = tss::stylesheet_archive::IdentifiedStyleEntry {
        identifier: style_identifier.to_owned(),
        style: reference,
    };
    let data = append_repeated_length_delimited_field(
        &stylesheet.messages[message_index].data,
        STYLESHEET_STYLES_FIELD,
        &reference.encode_to_vec(),
    )?;
    let data = append_repeated_length_delimited_field(
        &data,
        STYLESHEET_IDENTIFIER_MAP_FIELD,
        &entry.encode_to_vec(),
    )?;
    stylesheet.replace_message(
        message_index,
        RawMessage {
            type_: STYLESHEET_MESSAGE_TYPE,
            data,
        },
    )?;
    stylesheet.archive_info.message_infos[message_index]
        .object_references
        .push(style_id);
    Ok(archive.insert_object(style)?)
}

fn append_theme_preset(
    package: &mut IWorkPackage,
    location: &ThemeLocation,
    stylesheet_id: u64,
    style_id: u64,
) -> Result<()> {
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork theme {} is missing", location.object_id))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == location.message_type).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork theme {} must have exactly one theme payload",
                location.object_id
            )));
        };
        let message_index = *message_index;
        let mut theme = IWorkThemeArchive::decode(&object.messages[message_index].data)?;
        if theme
            .base
            .document_stylesheet
            .as_ref()
            .map(|reference| reference.identifier)
            != Some(stylesheet_id)
        {
            return Err(Error::InvalidFormat(format!(
                "iWork theme {} changed stylesheet during paragraph-style creation",
                location.object_id
            )));
        }
        let text = theme.extensions.text.as_mut().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork theme {} has no text presets",
                location.object_id
            ))
        })?;
        if text
            .paragraph_style_presets
            .iter()
            .any(|reference| reference.identifier == style_id)
        {
            return Err(Error::InvalidFormat(format!(
                "iWork theme {} already contains paragraph style {style_id}",
                location.object_id
            )));
        }
        text.paragraph_style_presets.push(tsp::Reference {
            identifier: style_id,
            ..Default::default()
        });
        object.replace_message(
            message_index,
            RawMessage {
                type_: location.message_type,
                data: theme.encode()?,
            },
        )?;
        object.archive_info.message_infos[message_index]
            .object_references
            .push(style_id);
        Ok(())
    })
}
