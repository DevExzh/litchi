//! Safe removal of unused named paragraph styles.

use prost::Message;

use crate::archive::RawMessage;
use crate::package_metadata::{
    component_identifier_for_entry, release_package_identifier_suffix,
    remove_component_external_references_to_object, remove_component_object_uuids,
};
use crate::protobuf::{tsp, tss};
use crate::wire::{repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields};
use crate::{Error, IWorkPackage, IWorkThemeArchive, Result};

use super::paragraph_alignment::native;
use super::paragraph_style_catalog::{ThemeLocation, locate_themes};
use super::style_registry::object_archive_name;
use litchi_iwa_text::paragraph::style::{NamedParagraphStyle, ParagraphStyleId, raw::native_id};

const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const STYLESHEET_IDENTIFIER_MAP_FIELD: u32 = 2;

pub(super) fn delete_named_paragraph_style(
    package: &mut IWorkPackage,
    first_style_id: u64,
    target: ParagraphStyleId,
) -> Result<NamedParagraphStyle> {
    native::validate_named_paragraph_style(package, first_style_id, target)?;
    let styles = native::named_paragraph_styles(package, first_style_id)?;
    if styles.len() <= 1 {
        return Err(Error::InvalidFormat(
            "iWork themes must retain at least one named paragraph style".to_owned(),
        ));
    }
    let deleted = styles
        .into_iter()
        .find(|style| style.id() == target)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {} is not selectable",
                native_id(target)
            ))
        })?;
    let location = native::locate_style(package, native_id(target))?;
    let stylesheet_id = native::stylesheet_id(&location.style, native_id(target))?;
    let stylesheet_archive_name = object_archive_name(package, stylesheet_id)?;
    if location.archive_name != stylesheet_archive_name {
        return Err(Error::InvalidFormat(format!(
            "named iWork paragraph style {} is not stored with stylesheet {stylesheet_id}",
            native_id(target)
        )));
    }
    let themes = locate_themes(package, stylesheet_id, native_id(target))?;
    ensure_no_live_references(package, native_id(target), stylesheet_id, &themes)?;

    let mut staged = package.clone();
    for theme in &themes {
        remove_theme_preset(&mut staged, theme, stylesheet_id, native_id(target))?;
    }
    remove_stylesheet_entry(
        &mut staged,
        &stylesheet_archive_name,
        stylesheet_id,
        native_id(target),
    )?;
    if let Some(component) = component_identifier_for_entry(&staged, &stylesheet_archive_name)? {
        remove_component_external_references_to_object(&mut staged, component, native_id(target))?;
        remove_component_object_uuids(&mut staged, component, &[native_id(target)])?;
    }
    release_package_identifier_suffix(&mut staged, &[native_id(target)])?;

    if native::named_paragraph_styles(&staged, first_style_id)?
        .iter()
        .any(|style| style.id() == target)
    {
        return Err(Error::InvalidFormat(
            "named iWork paragraph style deletion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(deleted)
}

fn ensure_no_live_references(
    package: &IWorkPackage,
    target: u64,
    stylesheet_id: u64,
    themes: &[ThemeLocation],
) -> Result<()> {
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        for object in &archive.objects {
            let object_id = object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork archive {archive_name} contains an object without an identifier"
                ))
            })?;
            for (message, info) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if !info.object_references.contains(&target) {
                    continue;
                }
                let is_stylesheet =
                    object_id == stylesheet_id && message.type_ == STYLESHEET_MESSAGE_TYPE;
                let is_theme = themes.iter().any(|theme| {
                    theme.archive_name == archive_name
                        && theme.object_id == object_id
                        && theme.message_type == message.type_
                });
                if !is_stylesheet && !is_theme {
                    return Err(Error::InvalidFormat(format!(
                        "iWork object {object_id} still references paragraph style {target}; choose or apply a replacement before deleting it"
                    )));
                }
            }
        }
    }
    Ok(())
}

fn remove_theme_preset(
    package: &mut IWorkPackage,
    location: &ThemeLocation,
    stylesheet_id: u64,
    target: u64,
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
                "iWork theme {} changed stylesheet during paragraph-style deletion",
                location.object_id
            )));
        }
        let text = theme.extensions.text.as_mut().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork theme {} has no text presets",
                location.object_id
            ))
        })?;
        ensure_only_paragraph_preset_reference(text, target)?;
        let occurrences = text
            .paragraph_style_presets
            .iter()
            .filter(|reference| reference.identifier == target)
            .count();
        if occurrences != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork theme {} references paragraph style {target} {occurrences} times",
                location.object_id
            )));
        }
        text.paragraph_style_presets
            .retain(|reference| reference.identifier != target);
        if text.paragraph_style_presets.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "iWork theme {} cannot lose its last paragraph style",
                location.object_id
            )));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: location.message_type,
                data: theme.encode()?,
            },
        )?;
        remove_metadata_reference(
            &mut object.archive_info.message_infos[message_index],
            target,
        );
        Ok(())
    })
}

fn ensure_only_paragraph_preset_reference(
    text: &crate::protobuf::tswp::ThemePresetsArchive,
    target: u64,
) -> Result<()> {
    let referenced_elsewhere = [
        text.list_style_presets.as_slice(),
        text.text_style_presets.as_slice(),
        text.imported_text_style_presets.as_slice(),
        text.toc_entry_style_presets.as_slice(),
        text.toc_settings_presets.as_slice(),
        text.character_style_presets.as_slice(),
        text.dropcap_style_presets.as_slice(),
    ]
    .into_iter()
    .flatten()
    .any(|reference| reference.identifier == target);
    if referenced_elsewhere {
        Err(Error::InvalidFormat(format!(
            "iWork text theme uses paragraph style {target} in another preset family"
        )))
    } else {
        Ok(())
    }
}

fn remove_stylesheet_entry(
    package: &mut IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    target: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let stylesheet = archive.object_mut(stylesheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork stylesheet {stylesheet_id} is missing"))
        })?;
        let indexes = stylesheet
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == STYLESHEET_MESSAGE_TYPE).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork stylesheet {stylesheet_id} must have exactly one stylesheet payload"
            )));
        };
        let message_index = *message_index;
        let original = &stylesheet.messages[message_index];
        let decoded = tss::StylesheetArchive::decode(original.data.as_slice())?;
        ensure_not_versioned_or_related(&decoded, target)?;
        let data = remove_main_style_entries(&original.data, target)?;
        stylesheet.replace_message(
            message_index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data,
            },
        )?;
        remove_metadata_reference(
            &mut stylesheet.archive_info.message_infos[message_index],
            target,
        );
        archive.remove_object(target).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork paragraph style {target} is missing"))
        })?;
        Ok(())
    })
}

fn remove_main_style_entries(data: &[u8], target: u64) -> Result<Vec<u8>> {
    let mut styles = repeated_length_delimited_payloads(data, STYLESHEET_STYLES_FIELD)?
        .into_iter()
        .map(|payload| {
            Ok((
                tsp::Reference::decode(payload)?.identifier,
                payload.to_vec(),
            ))
        })
        .collect::<Result<Vec<_>>>()?;
    let occurrences = styles.iter().filter(|(id, _)| *id == target).count();
    if occurrences != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet references paragraph style {target} {occurrences} times"
        )));
    }
    styles.retain(|(id, _)| *id != target);
    let style_payloads = styles
        .into_iter()
        .map(|(_, payload)| payload)
        .collect::<Vec<_>>();
    let data =
        rewrite_repeated_length_delimited_fields(data, STYLESHEET_STYLES_FIELD, &style_payloads)?;

    let mut entries = repeated_length_delimited_payloads(&data, STYLESHEET_IDENTIFIER_MAP_FIELD)?
        .into_iter()
        .map(|payload| {
            let entry = tss::stylesheet_archive::IdentifiedStyleEntry::decode(payload)?;
            Ok((entry.style.identifier, payload.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    let occurrences = entries.iter().filter(|(id, _)| *id == target).count();
    if occurrences != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet identifier map references paragraph style {target} {occurrences} times"
        )));
    }
    entries.retain(|(id, _)| *id != target);
    let entry_payloads = entries
        .into_iter()
        .map(|(_, payload)| payload)
        .collect::<Vec<_>>();
    rewrite_repeated_length_delimited_fields(
        &data,
        STYLESHEET_IDENTIFIER_MAP_FIELD,
        &entry_payloads,
    )
}

fn ensure_not_versioned_or_related(stylesheet: &tss::StylesheetArchive, target: u64) -> Result<()> {
    if stylesheet.parent_to_children_style_map.iter().any(|entry| {
        entry.parent.identifier == target
            || entry
                .children
                .iter()
                .any(|child| child.identifier == target)
    }) {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {target} participates in stylesheet inheritance"
        )));
    }
    let versioned = [
        stylesheet.styles_for_10_0.as_ref(),
        stylesheet.styles_for_10_1.as_ref(),
        stylesheet.styles_for_10_2.as_ref(),
        stylesheet.styles_for_11_0.as_ref(),
        stylesheet.styles_for_11_1.as_ref(),
        stylesheet.styles_for_11_2.as_ref(),
        stylesheet.styles_for_12_0.as_ref(),
        stylesheet.styles_for_12_1.as_ref(),
        stylesheet.styles_for_12_2.as_ref(),
        stylesheet.styles_for_13_0.as_ref(),
        stylesheet.styles_for_13_1.as_ref(),
        stylesheet.styles_for_13_2.as_ref(),
        stylesheet.styles_for_14_0.as_ref(),
        stylesheet.styles_for_14_1.as_ref(),
        stylesheet.styles_for_14_2.as_ref(),
        stylesheet.styles_for_14_4.as_ref(),
    ];
    if versioned.into_iter().flatten().any(|styles| {
        styles
            .styles
            .iter()
            .any(|reference| reference.identifier == target)
            || styles
                .identifier_to_style_map
                .iter()
                .any(|entry| entry.style.identifier == target)
            || styles.parent_to_children_style_map.iter().any(|entry| {
                entry.parent.identifier == target
                    || entry
                        .children
                        .iter()
                        .any(|child| child.identifier == target)
            })
    }) {
        return Err(Error::InvalidFormat(format!(
            "iWork paragraph style {target} participates in a versioned stylesheet registry"
        )));
    }
    Ok(())
}

fn remove_metadata_reference(info: &mut crate::archive::MessageInfo, target: u64) {
    info.object_references
        .retain(|reference| *reference != target);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|reference| *reference != target);
    }
}
