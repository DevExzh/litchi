//! Canonical native TSWP list-style objects.

use std::borrow::Cow;
use std::collections::HashSet;

use prost::Message;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::protobuf::tswp::list_style_archive::{LabelGeometry, LabelType, NumberType};
use crate::protobuf::{tsp, tss, tswp};
use crate::text::storage_wire::{locate_text_storages, update_parsed_archive};
use crate::wire::{
    patch_length_delimited_field, patch_varint_field, rewrite_repeated_fixed32_fields,
    rewrite_repeated_length_delimited_fields, rewrite_repeated_varint_fields,
};
use crate::{Error, IWorkPackage, Result};

use super::types::{
    ParagraphList, ParagraphListLabelColor, ParagraphListNumberFormat,
    ParagraphListNumberPunctuation, ParagraphListNumberSequence,
};

const LIST_STYLE_MESSAGE_TYPE: u32 = 2_023;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const LIST_LEVEL_COUNT: usize = 9;
const OVERRIDE_COUNT_FIELD: u32 = 10;
const TEXT_INDENTS_FIELD: u32 = 12;
const INDENTS_FIELD: u32 = 13;
const GEOMETRIES_FIELD: u32 = 14;
const NUMBER_TYPES_FIELD: u32 = 15;
const STRINGS_FIELD: u32 = 16;
const FONT_COLOR_NULL_FIELD: u32 = 20;
const FONT_COLOR_FIELD: u32 = 21;
const TIERED_NUMBERS_FIELD: u32 = 25;
const FONT_EM_POINTS: f32 = 11.0;
const NONE_INDENT_STEP_POINTS: f32 = 36.0;
const BULLET_INDENT_STEP_POINTS: f32 = 9.0;
const NUMBER_INDENT_STEP_POINTS: f32 = 18.0;
const BULLET_BASELINE_OFFSET_POINTS: f32 = -1.0;
const DEFAULT_LABEL_SCALE: f32 = 1.0;
const DOUBLE_PAREN_NUMBER_TYPE_OFFSET: i32 = 1;
const RIGHT_PAREN_NUMBER_TYPE_OFFSET: i32 = 2;
const BULLET_GLYPH: &str = "•";
const NONE_OVERRIDE_COUNT: u32 = 4;
const BULLET_OVERRIDE_COUNT: u32 = 5;
const NUMBER_OVERRIDE_COUNT: u32 = 6;

pub(super) struct ListStyleLocation {
    pub(super) object_id: u64,
    pub(super) archive_name: String,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) style: tswp::ListStyleArchive,
}

pub(super) struct LocatedListStyle {
    pub(super) location: ListStyleLocation,
    pub(super) archive: Archive,
    package_revision: u64,
}

pub(super) fn locate_style(package: &IWorkPackage, style_id: u64) -> Result<ListStyleLocation> {
    locate_style_with_archive(package, style_id).map(|located| located.location)
}

pub(super) fn locate_style_with_archive(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<LocatedListStyle> {
    let mut found = None;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(style_id) else {
            continue;
        };
        let payloads = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == LIST_STYLE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        let [(message_index, message)] = payloads.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} must have exactly one writable payload"
            )));
        };
        let style = tswp::ListStyleArchive::decode(message.data.as_slice())?;
        if found
            .replace(LocatedListStyle {
                location: ListStyleLocation {
                    object_id: style_id,
                    archive_name: archive_name.to_owned(),
                    message_index: *message_index,
                    message_type: message.type_,
                    style,
                },
                archive,
                package_revision: package.mutation_revision(),
            })
            .is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork list style {style_id} is missing")))
}

pub(super) fn stylesheet_id(
    package: &IWorkPackage,
    style: &tswp::ListStyleArchive,
    style_id: u64,
) -> Result<u64> {
    let mut current_id = style_id;
    let mut current = Cow::Borrowed(style);
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        if let Some(identifier) = current
            .super_
            .stylesheet
            .as_ref()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
        {
            return Ok(identifier);
        }
        current_id = current
            .super_
            .parent
            .as_ref()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {current_id} has neither a stylesheet nor a parent"
                ))
            })?;
        current = Cow::Owned(locate_style(package, current_id)?.style);
    }
}

pub(super) fn paragraph_list(style: &tswp::ListStyleArchive) -> Result<ParagraphList> {
    for preset in [
        ParagraphList::None,
        ParagraphList::Bullet,
        ParagraphList::Numbered,
    ] {
        if matches_preset(style, preset) {
            return Ok(preset);
        }
    }
    Err(Error::InvalidFormat(
        "iWork list style is not a supported canonical None, Bullet, or Numbered preset".to_owned(),
    ))
}

pub(super) fn resolved_paragraph_list(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<ParagraphList> {
    let mut current_id = style_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        let style = locate_style(package, current_id)?.style;
        if let Ok(preset) = paragraph_list(&style) {
            return Ok(preset);
        }
        current_id = parent_style_id(&style, current_id)?;
    }
}

pub(super) fn effective_bullet_strings(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<Vec<String>> {
    if resolved_paragraph_list(package, style_id)? != ParagraphList::Bullet {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not a text-bullet list"
        )));
    }
    let mut current_id = style_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        let style = locate_style(package, current_id)?.style;
        if !style.strings.is_empty() {
            validate_bullet_strings(current_id, &style.strings)?;
            return Ok(style.strings);
        }
        current_id = parent_style_id(&style, current_id)?;
    }
}

pub(super) fn effective_bullet_geometries(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<Vec<LabelGeometry>> {
    if resolved_paragraph_list(package, style_id)? != ParagraphList::Bullet {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not a text-bullet list"
        )));
    }
    effective_label_geometries(package, style_id)
}

pub(super) fn effective_label_geometries(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<Vec<LabelGeometry>> {
    if resolved_paragraph_list(package, style_id)? == ParagraphList::None {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} has no labels"
        )));
    }
    let mut current_id = style_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        let style = locate_style(package, current_id)?.style;
        if !style.geometries.is_empty() {
            validate_bullet_geometries(current_id, &style.geometries)?;
            return Ok(style.geometries);
        }
        current_id = parent_style_id(&style, current_id)?;
    }
}

pub(super) fn effective_list_indents(package: &IWorkPackage, style_id: u64) -> Result<Vec<f32>> {
    effective_list_float_array(package, style_id, ListFloatArray::LabelIndents)
}

pub(super) fn effective_list_text_indents(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<Vec<f32>> {
    effective_list_float_array(package, style_id, ListFloatArray::TextGaps)
}

pub(super) fn effective_label_color(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<ParagraphListLabelColor> {
    if resolved_paragraph_list(package, style_id)? == ParagraphList::None {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} does not have a label color"
        )));
    }
    let mut current_id = style_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        let style = locate_style(package, current_id)?.style;
        if style.font_color_null == Some(true) {
            return Ok(ParagraphListLabelColor::Automatic);
        }
        if let Some(color) = style.font_color.as_ref() {
            return Ok(ParagraphListLabelColor::Explicit(
                crate::shapes::color_from_native(color)?,
            ));
        }
        if style.font_color_null == Some(false) {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {current_id} enables a missing label color"
            )));
        }
        let Some(parent) = style
            .super_
            .parent
            .as_ref()
            .map(|parent| parent.identifier)
            .filter(|identifier| *identifier != 0)
        else {
            return Ok(ParagraphListLabelColor::Automatic);
        };
        current_id = parent;
    }
}

pub(super) fn effective_number_types(package: &IWorkPackage, style_id: u64) -> Result<Vec<i32>> {
    if resolved_paragraph_list(package, style_id)? != ParagraphList::Numbered {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not a numbered list"
        )));
    }
    let mut current_id = style_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        let style = locate_style(package, current_id)?.style;
        if !style.number_types.is_empty() {
            validate_number_types(current_id, &style.number_types)?;
            return Ok(style.number_types);
        }
        current_id = parent_style_id(&style, current_id)?;
    }
}

pub(super) fn effective_tiered_numbers(package: &IWorkPackage, style_id: u64) -> Result<Vec<bool>> {
    if resolved_paragraph_list(package, style_id)? != ParagraphList::Numbered {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} is not a numbered list"
        )));
    }
    let mut current_id = style_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        let style = locate_style(package, current_id)?.style;
        if !style.tiered_numbers.is_empty() {
            validate_tiered_numbers(current_id, &style.tiered_numbers)?;
            return Ok(style.tiered_numbers);
        }
        current_id = parent_style_id(&style, current_id)?;
    }
}

#[derive(Clone, Copy)]
enum ListFloatArray {
    LabelIndents,
    TextGaps,
}

fn effective_list_float_array(
    package: &IWorkPackage,
    style_id: u64,
    array: ListFloatArray,
) -> Result<Vec<f32>> {
    if resolved_paragraph_list(package, style_id)? == ParagraphList::None {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} does not have labels"
        )));
    }
    let mut current_id = style_id;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(current_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork list-style inheritance contains a cycle at {current_id}"
            )));
        }
        let style = locate_style(package, current_id)?.style;
        let values = match array {
            ListFloatArray::LabelIndents => &style.indents,
            ListFloatArray::TextGaps => &style.text_indents,
        };
        if !values.is_empty() {
            validate_list_float_array(current_id, values, array)?;
            return Ok(match array {
                ListFloatArray::LabelIndents => style.indents,
                ListFloatArray::TextGaps => style.text_indents,
            });
        }
        current_id = parent_style_id(&style, current_id)?;
    }
}

pub(super) fn parent_style_id(style: &tswp::ListStyleArchive, style_id: u64) -> Result<u64> {
    style
        .super_
        .parent
        .as_ref()
        .map(|parent| parent.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork list-style variation {style_id} has no parent"
            ))
        })
}

pub(super) fn variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    strings: Vec<String>,
) -> Result<ArchiveObject> {
    validate_bullet_strings(identifier, &strings)?;
    let style = tswp::ListStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        strings,
        ..Default::default()
    };
    let data = style.encode_to_vec();
    tswp::ListStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: LIST_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    Ok(object)
}

pub(super) fn geometry_variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    geometries: Vec<LabelGeometry>,
) -> Result<ArchiveObject> {
    validate_bullet_geometries(identifier, &geometries)?;
    let style = tswp::ListStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        geometries,
        ..Default::default()
    };
    let data = style.encode_to_vec();
    tswp::ListStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: LIST_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    Ok(object)
}

pub(super) fn indentation_variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    indents: Vec<f32>,
    text_indents: Vec<f32>,
) -> Result<ArchiveObject> {
    validate_list_float_array(identifier, &indents, ListFloatArray::LabelIndents)?;
    validate_list_float_array(identifier, &text_indents, ListFloatArray::TextGaps)?;
    let style = tswp::ListStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(2),
        text_indents,
        indents,
        ..Default::default()
    };
    let data = style.encode_to_vec();
    tswp::ListStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: LIST_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    Ok(object)
}

pub(super) fn label_color_variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    color: ParagraphListLabelColor,
) -> Result<ArchiveObject> {
    let (font_color_null, font_color) = native_label_color(color);
    let style = tswp::ListStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        font_color_null,
        font_color,
        ..Default::default()
    };
    let data = style.encode_to_vec();
    tswp::ListStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: LIST_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    Ok(object)
}

pub(super) fn number_format_variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    number_types: Vec<i32>,
) -> Result<ArchiveObject> {
    validate_number_types(identifier, &number_types)?;
    let style = tswp::ListStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        number_types,
        ..Default::default()
    };
    let data = style.encode_to_vec();
    tswp::ListStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: LIST_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    Ok(object)
}

pub(super) fn number_tiering_variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    tiered_numbers: Vec<bool>,
) -> Result<ArchiveObject> {
    validate_tiered_numbers(identifier, &tiered_numbers)?;
    let style = tswp::ListStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        tiered_numbers,
        ..Default::default()
    };
    let data = style.encode_to_vec();
    tswp::ListStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: LIST_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = STANDARD_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    Ok(object)
}

fn patch_style_message<F>(
    package: &mut IWorkPackage,
    location: &ListStyleLocation,
    patch: F,
) -> Result<()>
where
    F: FnOnce(&[u8], u64) -> Result<Option<Vec<u8>>>,
{
    let archive_name = location.archive_name.clone();
    package.update_archive(&archive_name, |archive| {
        patch_style_message_in_archive(archive, location, patch)
    })
}

fn patch_style_message_with_archive<F>(
    package: &mut IWorkPackage,
    located: LocatedListStyle,
    patch: F,
) -> Result<()>
where
    F: FnOnce(&[u8], u64) -> Result<Option<Vec<u8>>>,
{
    let LocatedListStyle {
        location,
        archive,
        package_revision,
    } = located;
    if package.mutation_revision() != package_revision {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {} package changed unexpectedly",
            location.object_id
        )));
    }
    let archive_name = location.archive_name.clone();
    update_parsed_archive(package, &archive_name, archive, |archive| {
        patch_style_message_in_archive(archive, &location, patch)
    })
}

fn patch_style_message_in_archive<F>(
    archive: &mut Archive,
    location: &ListStyleLocation,
    patch: F,
) -> Result<()>
where
    F: FnOnce(&[u8], u64) -> Result<Option<Vec<u8>>>,
{
    let style_id = location.object_id;
    let object = archive
        .object_mut(style_id)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork list style {style_id} is missing")))?;
    if object.archive_info.identifier != Some(style_id) {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} object identity changed unexpectedly"
        )));
    }
    if object.messages.get(location.message_index).is_none() {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} writable payload index {} is missing",
            location.message_index
        )));
    }
    if object.messages[location.message_index].type_ != location.message_type {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} writable payload changed unexpectedly"
        )));
    }
    if object
        .archive_info
        .message_infos
        .get(location.message_index)
        .is_none()
    {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} writable payload metadata index {} is missing",
            location.message_index
        )));
    }
    if object.archive_info.message_infos[location.message_index].type_ != location.message_type {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} writable payload metadata changed unexpectedly"
        )));
    }
    let replacement = {
        let original = &object.messages[location.message_index];
        patch(&original.data, style_id)?.map(|data| (original.type_, data))
    };
    let Some((message_type, data)) = replacement else {
        return Ok(());
    };
    object.replace_message(
        location.message_index,
        RawMessage {
            type_: message_type,
            data,
        },
    )?;
    Ok(())
}

pub(super) fn replace_direct_number_types_with_archive(
    package: &mut IWorkPackage,
    located: LocatedListStyle,
    number_types: &[i32],
) -> Result<()> {
    let style_id = located.location.object_id;
    validate_number_types(style_id, number_types)?;
    patch_style_message_with_archive(package, located, |data, style_id| {
        Ok(Some(patch_direct_number_types(
            data,
            style_id,
            number_types,
        )?))
    })
}

pub(super) fn remove_direct_number_types(
    package: &mut IWorkPackage,
    location: &ListStyleLocation,
) -> Result<()> {
    patch_style_message(package, location, |data, style_id| {
        let original = tswp::ListStyleArchive::decode(data)?;

        if original.number_types.is_empty() {
            return Ok(None);
        }
        validate_number_types(style_id, &original.number_types)?;
        let override_count = original
            .override_count
            .unwrap_or(0)
            .checked_sub(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} has number formats without an override count"
                ))
            })?;
        let data = rewrite_repeated_varint_fields(data, NUMBER_TYPES_FIELD, &[])?;
        let data = patch_varint_field(
            &data,
            OVERRIDE_COUNT_FIELD,
            original.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
        let decoded = tswp::ListStyleArchive::decode(data.as_slice())?;
        if !decoded.number_types.is_empty() || decoded.override_count != Some(override_count) {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} number-format removal failed validation"
            )));
        }
        Ok(Some(data))
    })
}

fn patch_direct_number_types(data: &[u8], style_id: u64, number_types: &[i32]) -> Result<Vec<u8>> {
    validate_number_types(style_id, number_types)?;
    let style = tswp::ListStyleArchive::decode(data)?;
    let had_direct_number_types = !style.number_types.is_empty();
    if had_direct_number_types {
        validate_number_types(style_id, &style.number_types)?;
    }
    let encoded = number_types
        .iter()
        .map(|value| u64::try_from(*value))
        .collect::<std::result::Result<Vec<_>, _>>()
        .map_err(|_| {
            Error::InvalidFormat(format!(
                "iWork list style {style_id} contains a negative number format"
            ))
        })?;
    let mut patched = rewrite_repeated_varint_fields(data, NUMBER_TYPES_FIELD, &encoded)?;
    if !had_direct_number_types {
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} override count overflowed"
                ))
            })?;
        patched = patch_varint_field(
            &patched,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
    }
    let decoded = tswp::ListStyleArchive::decode(patched.as_slice())?;
    if decoded.number_types != number_types {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} number-format update failed validation"
        )));
    }
    Ok(patched)
}

pub(super) fn replace_direct_tiered_numbers_with_archive(
    package: &mut IWorkPackage,
    located: LocatedListStyle,
    tiered_numbers: &[bool],
) -> Result<()> {
    let style_id = located.location.object_id;
    validate_tiered_numbers(style_id, tiered_numbers)?;
    patch_style_message_with_archive(package, located, |data, style_id| {
        Ok(Some(patch_direct_tiered_numbers(
            data,
            style_id,
            tiered_numbers,
        )?))
    })
}

pub(super) fn remove_direct_tiered_numbers(
    package: &mut IWorkPackage,
    location: &ListStyleLocation,
) -> Result<()> {
    patch_style_message(package, location, |data, style_id| {
        let style = tswp::ListStyleArchive::decode(data)?;
        if style.tiered_numbers.is_empty() {
            return Ok(None);
        }
        validate_tiered_numbers(style_id, &style.tiered_numbers)?;
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_sub(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} has number tiering without an override count"
                ))
            })?;
        let data = rewrite_repeated_varint_fields(data, TIERED_NUMBERS_FIELD, &[])?;
        let data = patch_varint_field(
            &data,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
        let decoded = tswp::ListStyleArchive::decode(data.as_slice())?;
        if !decoded.tiered_numbers.is_empty() || decoded.override_count != Some(override_count) {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} number-tiering removal failed validation"
            )));
        }
        Ok(Some(data))
    })
}

fn patch_direct_tiered_numbers(
    data: &[u8],
    style_id: u64,
    tiered_numbers: &[bool],
) -> Result<Vec<u8>> {
    validate_tiered_numbers(style_id, tiered_numbers)?;
    let style = tswp::ListStyleArchive::decode(data)?;
    let had_direct_tiering = !style.tiered_numbers.is_empty();
    if had_direct_tiering {
        validate_tiered_numbers(style_id, &style.tiered_numbers)?;
    }
    let encoded = tiered_numbers
        .iter()
        .copied()
        .map(u64::from)
        .collect::<Vec<_>>();
    let mut patched = rewrite_repeated_varint_fields(data, TIERED_NUMBERS_FIELD, &encoded)?;
    if !had_direct_tiering {
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} override count overflowed"
                ))
            })?;
        patched = patch_varint_field(
            &patched,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
    }
    let decoded = tswp::ListStyleArchive::decode(patched.as_slice())?;
    if decoded.tiered_numbers != tiered_numbers {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} number-tiering update failed validation"
        )));
    }
    Ok(patched)
}

pub(super) fn replace_direct_label_color_with_archive(
    package: &mut IWorkPackage,
    located: LocatedListStyle,
    color: ParagraphListLabelColor,
) -> Result<()> {
    patch_style_message_with_archive(package, located, |data, style_id| {
        Ok(Some(patch_direct_label_color(data, style_id, color)?))
    })
}

pub(super) fn remove_direct_label_color(
    package: &mut IWorkPackage,
    location: &ListStyleLocation,
) -> Result<()> {
    patch_style_message(package, location, |data, style_id| {
        let style = tswp::ListStyleArchive::decode(data)?;
        let had_override = style.font_color_null.is_some() || style.font_color.is_some();
        if !had_override {
            return Ok(None);
        }
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_sub(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} has label color without an override count"
                ))
            })?;
        let data = patch_varint_field(
            data,
            FONT_COLOR_NULL_FIELD,
            style.font_color_null.is_some(),
            None,
        )?;
        let data = patch_length_delimited_field(
            &data,
            FONT_COLOR_FIELD,
            style.font_color.is_some(),
            None,
        )?;
        let data = patch_varint_field(
            &data,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
        let decoded = tswp::ListStyleArchive::decode(data.as_slice())?;
        if decoded.font_color_null.is_some()
            || decoded.font_color.is_some()
            || decoded.override_count != Some(override_count)
        {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} label-color removal failed validation"
            )));
        }
        Ok(Some(data))
    })
}

fn patch_direct_label_color(
    data: &[u8],
    style_id: u64,
    color: ParagraphListLabelColor,
) -> Result<Vec<u8>> {
    let style = tswp::ListStyleArchive::decode(data)?;
    let had_override = style.font_color_null.is_some() || style.font_color.is_some();
    let (font_color_null, font_color) = native_label_color(color);
    let mut patched = patch_varint_field(
        data,
        FONT_COLOR_NULL_FIELD,
        style.font_color_null.is_some(),
        font_color_null.map(u64::from),
    )?;
    let encoded_color = font_color.as_ref().map(Message::encode_to_vec);
    patched = patch_length_delimited_field(
        &patched,
        FONT_COLOR_FIELD,
        style.font_color.is_some(),
        encoded_color.as_deref(),
    )?;
    if !had_override {
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} override count overflowed"
                ))
            })?;
        patched = patch_varint_field(
            &patched,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
    }
    let decoded = tswp::ListStyleArchive::decode(patched.as_slice())?;
    if decoded.font_color_null != font_color_null || decoded.font_color != font_color {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} label-color update failed validation"
        )));
    }
    Ok(patched)
}

fn native_label_color(color: ParagraphListLabelColor) -> (Option<bool>, Option<tsp::Color>) {
    match color {
        ParagraphListLabelColor::Automatic => (Some(true), None),
        ParagraphListLabelColor::Explicit(color) => {
            (None, Some(crate::shapes::color_to_native(color)))
        },
    }
}

pub(super) fn replace_direct_bullet_strings_with_archive(
    package: &mut IWorkPackage,
    located: LocatedListStyle,
    strings: &[String],
) -> Result<()> {
    let style_id = located.location.object_id;
    validate_bullet_strings(style_id, strings)?;
    patch_style_message_with_archive(package, located, |data, style_id| {
        Ok(Some(patch_direct_bullet_strings(data, style_id, strings)?))
    })
}

pub(super) fn replace_direct_bullet_geometries_with_archive(
    package: &mut IWorkPackage,
    located: LocatedListStyle,
    geometries: &[LabelGeometry],
) -> Result<()> {
    let style_id = located.location.object_id;
    validate_bullet_geometries(style_id, geometries)?;
    patch_style_message_with_archive(package, located, |data, style_id| {
        Ok(Some(patch_direct_bullet_geometries(
            data, style_id, geometries,
        )?))
    })
}

pub(super) fn remove_direct_bullet_geometries(
    package: &mut IWorkPackage,
    location: &ListStyleLocation,
) -> Result<()> {
    patch_style_message(package, location, |data, style_id| {
        let style = tswp::ListStyleArchive::decode(data)?;
        if style.geometries.is_empty() {
            return Ok(None);
        }
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_sub(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} has geometry without an override count"
                ))
            })?;
        let data = rewrite_repeated_length_delimited_fields(data, GEOMETRIES_FIELD, &[])?;
        let data = patch_varint_field(
            &data,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
        let decoded = tswp::ListStyleArchive::decode(data.as_slice())?;
        if !decoded.geometries.is_empty() || decoded.override_count != Some(override_count) {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} bullet geometry removal failed validation"
            )));
        }
        Ok(Some(data))
    })
}

fn patch_direct_bullet_geometries(
    data: &[u8],
    style_id: u64,
    geometries: &[LabelGeometry],
) -> Result<Vec<u8>> {
    validate_bullet_geometries(style_id, geometries)?;
    let style = tswp::ListStyleArchive::decode(data)?;
    let had_direct_geometries = !style.geometries.is_empty();
    if had_direct_geometries {
        validate_bullet_geometries(style_id, &style.geometries)?;
    }
    let encoded = geometries
        .iter()
        .map(Message::encode_to_vec)
        .collect::<Vec<_>>();
    let mut patched = rewrite_repeated_length_delimited_fields(data, GEOMETRIES_FIELD, &encoded)?;
    if !had_direct_geometries {
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} override count overflowed"
                ))
            })?;
        patched = patch_varint_field(
            &patched,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
    }
    let decoded = tswp::ListStyleArchive::decode(patched.as_slice())?;
    if decoded.geometries != geometries {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} bullet geometry update failed validation"
        )));
    }
    Ok(patched)
}

pub(super) fn replace_direct_list_indentation_with_archive(
    package: &mut IWorkPackage,
    located: LocatedListStyle,
    indents: &[f32],
    text_indents: &[f32],
) -> Result<()> {
    let style_id = located.location.object_id;
    validate_list_float_array(style_id, indents, ListFloatArray::LabelIndents)?;
    validate_list_float_array(style_id, text_indents, ListFloatArray::TextGaps)?;
    patch_style_message_with_archive(package, located, |data, style_id| {
        Ok(Some(patch_direct_list_indentation(
            data,
            style_id,
            indents,
            text_indents,
        )?))
    })
}

fn patch_direct_list_indentation(
    data: &[u8],
    style_id: u64,
    indents: &[f32],
    text_indents: &[f32],
) -> Result<Vec<u8>> {
    validate_list_float_array(style_id, indents, ListFloatArray::LabelIndents)?;
    validate_list_float_array(style_id, text_indents, ListFloatArray::TextGaps)?;
    let style = tswp::ListStyleArchive::decode(data)?;
    if !style.indents.is_empty() {
        validate_list_float_array(style_id, &style.indents, ListFloatArray::LabelIndents)?;
    }
    if !style.text_indents.is_empty() {
        validate_list_float_array(style_id, &style.text_indents, ListFloatArray::TextGaps)?;
    }
    let missing_overrides =
        u32::from(style.indents.is_empty()) + u32::from(style.text_indents.is_empty());
    let indent_bits = indents
        .iter()
        .map(|value| value.to_bits())
        .collect::<Vec<_>>();
    let text_indent_bits = text_indents
        .iter()
        .map(|value| value.to_bits())
        .collect::<Vec<_>>();
    let data = rewrite_repeated_fixed32_fields(data, INDENTS_FIELD, &indent_bits)?;
    let mut data = rewrite_repeated_fixed32_fields(&data, TEXT_INDENTS_FIELD, &text_indent_bits)?;
    if missing_overrides != 0 {
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_add(missing_overrides)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} override count overflowed"
                ))
            })?;
        data = patch_varint_field(
            &data,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
    }
    let decoded = tswp::ListStyleArchive::decode(data.as_slice())?;
    if decoded.indents != indents || decoded.text_indents != text_indents {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} indentation update failed validation"
        )));
    }
    Ok(data)
}

pub(super) fn remove_direct_list_indentation(
    package: &mut IWorkPackage,
    location: &ListStyleLocation,
) -> Result<()> {
    patch_style_message(package, location, |data, style_id| {
        let style = tswp::ListStyleArchive::decode(data)?;
        let removed_overrides =
            u32::from(!style.indents.is_empty()) + u32::from(!style.text_indents.is_empty());
        if removed_overrides == 0 {
            return Ok(None);
        }
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_sub(removed_overrides)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} has indentation without enough overrides"
                ))
            })?;
        let data = rewrite_repeated_fixed32_fields(data, INDENTS_FIELD, &[])?;
        let data = rewrite_repeated_fixed32_fields(&data, TEXT_INDENTS_FIELD, &[])?;
        let data = patch_varint_field(
            &data,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
        let decoded = tswp::ListStyleArchive::decode(data.as_slice())?;
        if !decoded.indents.is_empty()
            || !decoded.text_indents.is_empty()
            || decoded.override_count != Some(override_count)
        {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} indentation removal failed validation"
            )));
        }
        Ok(Some(data))
    })
}

fn patch_direct_bullet_strings(data: &[u8], style_id: u64, strings: &[String]) -> Result<Vec<u8>> {
    validate_bullet_strings(style_id, strings)?;
    let style = tswp::ListStyleArchive::decode(data)?;
    let had_direct_strings = !style.strings.is_empty();
    if had_direct_strings {
        validate_bullet_strings(style_id, &style.strings)?;
    }
    let encoded = strings
        .iter()
        .map(|string| string.as_bytes().to_vec())
        .collect::<Vec<_>>();
    let mut patched = rewrite_repeated_length_delimited_fields(data, STRINGS_FIELD, &encoded)?;
    if !had_direct_strings {
        let override_count = style
            .override_count
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork list style {style_id} override count overflowed"
                ))
            })?;
        patched = patch_varint_field(
            &patched,
            OVERRIDE_COUNT_FIELD,
            style.override_count.is_some(),
            Some(u64::from(override_count)),
        )?;
    }
    let decoded = tswp::ListStyleArchive::decode(patched.as_slice())?;
    if decoded.strings != strings {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} bullet update failed validation"
        )));
    }
    Ok(patched)
}

pub(super) fn is_exclusive(package: &IWorkPackage, style_id: u64) -> Result<bool> {
    let storage_references = locate_text_storages(package)?
        .into_iter()
        .map(|location| {
            location
                .storage
                .table_list_style
                .iter()
                .flat_map(|table| &table.entries)
                .filter(|entry| {
                    entry
                        .object
                        .as_ref()
                        .is_some_and(|reference| reference.identifier == style_id)
                })
                .count()
        })
        .sum::<usize>();

    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            for message in &object.messages {
                if message.type_ == LIST_STYLE_MESSAGE_TYPE && {
                    let style = tswp::ListStyleArchive::decode(message.data.as_slice())?;
                    style
                        .super_
                        .parent
                        .as_ref()
                        .is_some_and(|parent| parent.identifier == style_id)
                } {
                    return Ok(false);
                }
            }
        }
    }
    Ok(storage_references == 1)
}

fn validate_bullet_strings(style_id: u64, strings: &[String]) -> Result<()> {
    if strings.len() != LIST_LEVEL_COUNT {
        return Err(Error::InvalidFormat(format!(
            "iWork text-bullet list style {style_id} must define {LIST_LEVEL_COUNT} levels, found {}",
            strings.len()
        )));
    }
    if strings.iter().any(String::is_empty) {
        return Err(Error::InvalidFormat(format!(
            "iWork text-bullet list style {style_id} contains an empty marker"
        )));
    }
    Ok(())
}

fn validate_bullet_geometries(style_id: u64, geometries: &[LabelGeometry]) -> Result<()> {
    if geometries.len() != LIST_LEVEL_COUNT {
        return Err(Error::InvalidFormat(format!(
            "iWork text-bullet list style {style_id} must define {LIST_LEVEL_COUNT} geometries, found {}",
            geometries.len()
        )));
    }
    for geometry in geometries {
        let scale = geometry.scale.unwrap_or(DEFAULT_LABEL_SCALE);
        let baseline = geometry.baseline_offset.unwrap_or(0.0);
        if !scale.is_finite() || scale <= 0.0 || !baseline.is_finite() {
            return Err(Error::InvalidFormat(format!(
                "iWork text-bullet list style {style_id} contains invalid geometry"
            )));
        }
        if geometry.scale_with_text == Some(false) {
            return Err(Error::InvalidFormat(format!(
                "iWork text-bullet list style {style_id} contains an absolute-size geometry"
            )));
        }
    }
    Ok(())
}

fn validate_list_float_array(style_id: u64, values: &[f32], array: ListFloatArray) -> Result<()> {
    if values.len() != LIST_LEVEL_COUNT {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} must define {LIST_LEVEL_COUNT} {} values, found {}",
            match array {
                ListFloatArray::LabelIndents => "label-indent",
                ListFloatArray::TextGaps => "text-gap",
            },
            values.len()
        )));
    }
    if values
        .iter()
        .any(|value| !value.is_finite() || *value < 0.0)
    {
        return Err(Error::InvalidFormat(format!(
            "iWork list style {style_id} contains an invalid {} value",
            match array {
                ListFloatArray::LabelIndents => "label-indent",
                ListFloatArray::TextGaps => "text-gap",
            }
        )));
    }
    Ok(())
}

fn validate_number_types(style_id: u64, values: &[i32]) -> Result<()> {
    if values.len() != LIST_LEVEL_COUNT {
        return Err(Error::InvalidFormat(format!(
            "iWork numbered list style {style_id} must define {LIST_LEVEL_COUNT} number formats, found {}",
            values.len()
        )));
    }
    for value in values {
        number_format_from_native(*value)?;
    }
    Ok(())
}

fn validate_tiered_numbers(style_id: u64, values: &[bool]) -> Result<()> {
    if values.len() != LIST_LEVEL_COUNT {
        return Err(Error::InvalidFormat(format!(
            "iWork numbered list style {style_id} must define {LIST_LEVEL_COUNT} tiering values, found {}",
            values.len()
        )));
    }
    Ok(())
}

pub(super) fn number_format_to_native(format: ParagraphListNumberFormat) -> i32 {
    match format {
        ParagraphListNumberFormat::Circled => NumberType::KCircledNumberKind as i32,
        ParagraphListNumberFormat::HebrewBiblicalStandard => {
            NumberType::KHebrewBiblicalStandardKind as i32
        },
        ParagraphListNumberFormat::Affixed {
            sequence,
            punctuation,
        } => {
            let base = match sequence {
                ParagraphListNumberSequence::Decimal => NumberType::KNumericDecimal,
                ParagraphListNumberSequence::RomanUppercase => NumberType::KRomanUpperDecimal,
                ParagraphListNumberSequence::RomanLowercase => NumberType::KRomanLowerDecimal,
                ParagraphListNumberSequence::LatinUppercase => NumberType::KAlphaUpperDecimal,
                ParagraphListNumberSequence::LatinLowercase => NumberType::KAlphaLowerDecimal,
                ParagraphListNumberSequence::JapaneseIdeographic => {
                    NumberType::KIdeographicJapaneseDecimalKind
                },
                ParagraphListNumberSequence::JapaneseHiragana => NumberType::KHiraganaDecimalKind,
                ParagraphListNumberSequence::JapaneseKatakana => NumberType::KKatakanaDecimalKind,
                ParagraphListNumberSequence::JapaneseHiraganaIroha => {
                    NumberType::KHiraganaIrohaDecimalKind
                },
                ParagraphListNumberSequence::JapaneseKatakanaIroha => {
                    NumberType::KKatakanaIrohaDecimalKind
                },
                ParagraphListNumberSequence::SimplifiedChineseIdeographic => {
                    NumberType::KIdeographicSimplifiedChineseDecimalKind
                },
                ParagraphListNumberSequence::TraditionalChineseIdeographic => {
                    NumberType::KIdeographicTraditionalChineseDecimalKind
                },
                ParagraphListNumberSequence::FormalJapaneseIdeographic => {
                    NumberType::KIdeographicFormalJapaneseDecimalKind
                },
                ParagraphListNumberSequence::FormalSimplifiedChineseIdeographic => {
                    NumberType::KIdeographicFormalSimplifiedChineseDecimalKind
                },
                ParagraphListNumberSequence::FormalTraditionalChineseIdeographic => {
                    NumberType::KIdeographicFormalTraditionalChineseDecimalKind
                },
                ParagraphListNumberSequence::KoreanAlphabet => {
                    NumberType::KKoreanAlphabetDecimalKind
                },
                ParagraphListNumberSequence::ArabicIndic => NumberType::KArabianNumericDecimalKind,
                ParagraphListNumberSequence::ArabicAlphabet => NumberType::KArabianAlphaDecimalKind,
                ParagraphListNumberSequence::ArabicAbjad => NumberType::KArabianAbjadDecimalKind,
                ParagraphListNumberSequence::HebrewAlphabet => NumberType::KHebrewAlphaDecimalKind,
                ParagraphListNumberSequence::HebrewBiblical => {
                    NumberType::KHebrewBiblicalDecimalKind
                },
            } as i32;
            base + match punctuation {
                ParagraphListNumberPunctuation::Period => 0,
                ParagraphListNumberPunctuation::Parentheses => DOUBLE_PAREN_NUMBER_TYPE_OFFSET,
                ParagraphListNumberPunctuation::RightParenthesis => RIGHT_PAREN_NUMBER_TYPE_OFFSET,
            }
        },
    }
}

pub(super) fn number_format_from_native(value: i32) -> Result<ParagraphListNumberFormat> {
    NumberType::try_from(value).map_err(|_| {
        Error::InvalidFormat(format!(
            "native iWork numbered-list format {value} is unknown"
        ))
    })?;
    if value == NumberType::KCircledNumberKind as i32 {
        return Ok(ParagraphListNumberFormat::Circled);
    }
    if value == NumberType::KHebrewBiblicalStandardKind as i32 {
        return Ok(ParagraphListNumberFormat::HebrewBiblicalStandard);
    }
    for sequence in ParagraphListNumberSequence::ALL {
        for punctuation in ParagraphListNumberPunctuation::ALL {
            let format = ParagraphListNumberFormat::affixed(sequence, punctuation);
            if number_format_to_native(format) == value {
                return Ok(format);
            }
        }
    }
    Err(Error::InvalidFormat(format!(
        "native iWork numbered-list format {value} has no public representation"
    )))
}

pub(super) fn find_preset_style(
    package: &IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    preset: ParagraphList,
) -> Result<Option<u64>> {
    let archive = package.archive(archive_name)?;
    find_preset_style_in_archive(&archive, stylesheet_id, preset)
}

pub(super) fn find_preset_style_in_archive(
    archive: &Archive,
    stylesheet_id: u64,
    preset: ParagraphList,
) -> Result<Option<u64>> {
    let mut identifiers = Vec::new();
    for object in &archive.objects {
        let Some(identifier) = object.archive_info.identifier else {
            continue;
        };
        for message in &object.messages {
            if message.type_ != LIST_STYLE_MESSAGE_TYPE {
                continue;
            }
            let style = tswp::ListStyleArchive::decode(message.data.as_slice())?;
            if style
                .super_
                .stylesheet
                .as_ref()
                .is_some_and(|reference| reference.identifier == stylesheet_id)
                && matches_preset(&style, preset)
            {
                identifiers.push(identifier);
            }
        }
    }
    identifiers.sort_unstable();
    Ok(identifiers.into_iter().next())
}

pub(super) fn style_object(
    identifier: u64,
    stylesheet_id: u64,
    preset: ParagraphList,
) -> Result<ArchiveObject> {
    let mut style = canonical_archive(preset);
    style.super_ = tss::StyleArchive {
        name: Some(preset.native_name().to_owned()),
        stylesheet: Some(reference(stylesheet_id)),
        ..Default::default()
    };
    let data = style.encode_to_vec();
    tswp::ListStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: LIST_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    Ok(object)
}

fn matches_preset(style: &tswp::ListStyleArchive, preset: ParagraphList) -> bool {
    let expected = canonical_archive(preset);
    if preset == ParagraphList::None
        && style.super_.is_variation != Some(true)
        && (style.label_types.is_empty()
            || style
                .label_types
                .iter()
                .all(|label| *label == LabelType::KNone as i32))
    {
        return style.strings.is_empty()
            && style.number_types.is_empty()
            && style.images.is_empty();
    }
    style.override_count == expected.override_count
        && style.label_types == expected.label_types
        && style.text_indents == expected.text_indents
        && style.indents == expected.indents
        && style.geometries == expected.geometries
        && style.number_types == expected.number_types
        && style.strings == expected.strings
        && style.images == expected.images
        && style.shadow_null == expected.shadow_null
        && style.shadow == expected.shadow
        && style.font_color_null == expected.font_color_null
        && style.font_color == expected.font_color
        && style.font_name_null == expected.font_name_null
        && style.font_name == expected.font_name
        && style.writing_direction == expected.writing_direction
        && style.tiered_numbers == expected.tiered_numbers
}

fn canonical_archive(preset: ParagraphList) -> tswp::ListStyleArchive {
    match preset {
        ParagraphList::None => tswp::ListStyleArchive {
            super_: tss::StyleArchive::default(),
            override_count: Some(NONE_OVERRIDE_COUNT),
            label_types: repeated_enum(LabelType::KNone),
            text_indents: vec![0.0; LIST_LEVEL_COUNT],
            indents: level_indents(NONE_INDENT_STEP_POINTS),
            geometries: repeated_geometry(|_| 0.0),
            ..Default::default()
        },
        ParagraphList::Bullet => tswp::ListStyleArchive {
            super_: tss::StyleArchive::default(),
            override_count: Some(BULLET_OVERRIDE_COUNT),
            label_types: repeated_enum(LabelType::KString),
            text_indents: vec![BULLET_INDENT_STEP_POINTS / FONT_EM_POINTS; LIST_LEVEL_COUNT],
            indents: level_indents(BULLET_INDENT_STEP_POINTS),
            geometries: repeated_geometry(|level| {
                if level == 0 {
                    0.0
                } else {
                    BULLET_BASELINE_OFFSET_POINTS
                }
            }),
            strings: vec![BULLET_GLYPH.to_owned(); LIST_LEVEL_COUNT],
            ..Default::default()
        },
        ParagraphList::Numbered => tswp::ListStyleArchive {
            super_: tss::StyleArchive::default(),
            override_count: Some(NUMBER_OVERRIDE_COUNT),
            label_types: repeated_enum(LabelType::KNumber),
            text_indents: vec![NUMBER_INDENT_STEP_POINTS / FONT_EM_POINTS; LIST_LEVEL_COUNT],
            indents: level_indents(NUMBER_INDENT_STEP_POINTS),
            geometries: repeated_geometry(|_| 0.0),
            number_types: repeated_enum(NumberType::KNumericDecimal),
            tiered_numbers: vec![false; LIST_LEVEL_COUNT],
            ..Default::default()
        },
    }
}

fn repeated_enum<T>(value: T) -> Vec<i32>
where
    T: Copy + Into<i32>,
{
    vec![value.into(); LIST_LEVEL_COUNT]
}

fn level_indents(step: f32) -> Vec<f32> {
    (0..LIST_LEVEL_COUNT)
        .map(|level| level as f32 * step)
        .collect()
}

fn repeated_geometry(baseline: impl Fn(usize) -> f32) -> Vec<LabelGeometry> {
    (0..LIST_LEVEL_COUNT)
        .map(|level| LabelGeometry {
            scale: Some(DEFAULT_LABEL_SCALE),
            baseline_offset: Some(baseline(level)),
            scale_with_text: Some(true),
        })
        .collect()
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

    #[test]
    fn canonical_presets_are_strict_and_reversible() {
        for preset in [
            ParagraphList::None,
            ParagraphList::Bullet,
            ParagraphList::Numbered,
        ] {
            let archive = canonical_archive(preset);
            assert_eq!(paragraph_list(&archive).unwrap(), preset);
            assert_eq!(archive.label_types.len(), LIST_LEVEL_COUNT);
            assert_eq!(archive.text_indents.len(), LIST_LEVEL_COUNT);
            assert_eq!(archive.indents.len(), LIST_LEVEL_COUNT);
            assert_eq!(archive.geometries.len(), LIST_LEVEL_COUNT);
        }

        let mut invalid = canonical_archive(ParagraphList::Bullet);
        invalid.strings[0] = "-".to_owned();
        assert!(paragraph_list(&invalid).is_err());
    }

    #[test]
    fn preset_lookup_reuses_one_parsed_archive() {
        let archive = Archive {
            objects: vec![style_object(10, 3, ParagraphList::Bullet).unwrap()],
        };

        assert_eq!(
            find_preset_style_in_archive(&archive, 3, ParagraphList::Bullet).unwrap(),
            Some(10)
        );
        assert_eq!(
            find_preset_style_in_archive(&archive, 3, ParagraphList::Numbered).unwrap(),
            None
        );
    }

    #[test]
    fn bullet_variations_encode_all_nine_levels() {
        let mut strings = vec![BULLET_GLYPH.to_owned(); LIST_LEVEL_COUNT];
        strings[1] = "➡".to_owned();
        let object = variation_object(10, 8, 3, strings.clone()).unwrap();
        let style = tswp::ListStyleArchive::decode(object.messages[0].data.as_slice()).unwrap();
        assert_eq!(style.override_count, Some(1));
        assert_eq!(style.strings, strings);
        assert_eq!(style.super_.parent.unwrap().identifier, 8);
        assert_eq!(style.super_.stylesheet.unwrap().identifier, 3);
        assert_eq!(object.archive_info.message_infos[0].object_references, [8]);
    }

    #[test]
    fn bullet_updates_preserve_unknown_wire_fields() {
        let mut style = canonical_archive(ParagraphList::Bullet).encode_to_vec();
        let unknown = [0x98, 0x06, 0x07];
        style.extend_from_slice(&unknown);
        let mut strings = vec![BULLET_GLYPH.to_owned(); LIST_LEVEL_COUNT];
        strings[3] = "◆".to_owned();
        let patched = patch_direct_bullet_strings(&style, 10, &strings).unwrap();
        assert!(
            patched
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        assert_eq!(
            tswp::ListStyleArchive::decode(patched.as_slice())
                .unwrap()
                .strings,
            strings
        );
    }

    #[test]
    fn exclusivity_rejects_malformed_recognized_storage() {
        let object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2_001,
                data: vec![0x80],
            }],
        )
        .unwrap();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/One.iwa",
                &crate::archive::Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        assert!(is_exclusive(&package, 42).is_err());
    }

    #[test]
    fn geometry_updates_compose_with_glyph_overrides_losslessly() {
        let mut style = tswp::ListStyleArchive {
            super_: tss::StyleArchive {
                parent: Some(reference(8)),
                is_variation: Some(true),
                stylesheet: Some(reference(3)),
                ..Default::default()
            },
            override_count: Some(1),
            strings: vec![BULLET_GLYPH.to_owned(); LIST_LEVEL_COUNT],
            ..Default::default()
        };
        style.strings[1] = "➡".to_owned();
        let mut encoded = style.encode_to_vec();
        let unknown = [0xa0, 0x06, 0x07];
        encoded.extend_from_slice(&unknown);

        let mut geometries = repeated_geometry(|level| {
            if level == 0 {
                0.0
            } else {
                BULLET_BASELINE_OFFSET_POINTS
            }
        });
        geometries[1].scale = Some(1.75);
        geometries[1].baseline_offset = Some(4.0);
        let patched = patch_direct_bullet_geometries(&encoded, 10, &geometries).unwrap();
        assert!(
            patched
                .windows(unknown.len())
                .any(|window| window == unknown)
        );

        let decoded = tswp::ListStyleArchive::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.override_count, Some(2));
        assert_eq!(decoded.strings[1], "➡");
        assert_eq!(decoded.geometries, geometries);
    }

    #[test]
    fn indentation_updates_compose_with_glyph_and_geometry_losslessly() {
        let mut style = tswp::ListStyleArchive {
            super_: tss::StyleArchive {
                parent: Some(reference(8)),
                is_variation: Some(true),
                stylesheet: Some(reference(3)),
                ..Default::default()
            },
            override_count: Some(2),
            geometries: repeated_geometry(|level| {
                if level == 0 {
                    0.0
                } else {
                    BULLET_BASELINE_OFFSET_POINTS
                }
            }),
            strings: vec![BULLET_GLYPH.to_owned(); LIST_LEVEL_COUNT],
            ..Default::default()
        };
        style.strings[1] = "➡".to_owned();
        style.geometries[1].scale = Some(1.5);
        style.geometries[1].baseline_offset = Some(2.0);
        let mut encoded = style.encode_to_vec();
        let unknown = [0xa8, 0x06, 0x09];
        encoded.extend_from_slice(&unknown);
        let mut indents = level_indents(BULLET_INDENT_STEP_POINTS);
        let mut text_indents = vec![BULLET_INDENT_STEP_POINTS / FONT_EM_POINTS; LIST_LEVEL_COUNT];
        indents[1] = 10.0;
        text_indents[1] = 10.0 / 12.0;

        let patched = patch_direct_list_indentation(&encoded, 10, &indents, &text_indents).unwrap();
        assert!(
            patched
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        let decoded = tswp::ListStyleArchive::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.override_count, Some(4));
        assert_eq!(decoded.strings[1], "➡");
        assert_eq!(decoded.geometries[1].scale, Some(1.5));
        assert_eq!(decoded.indents, indents);
        assert_eq!(decoded.text_indents, text_indents);
    }

    #[test]
    fn label_color_updates_compose_with_other_overrides_and_preserve_unknown_fields() {
        let style = tswp::ListStyleArchive {
            super_: tss::StyleArchive {
                parent: Some(reference(8)),
                is_variation: Some(true),
                stylesheet: Some(reference(3)),
                ..Default::default()
            },
            override_count: Some(4),
            geometries: repeated_geometry(|_| 0.0),
            strings: vec![BULLET_GLYPH.to_owned(); LIST_LEVEL_COUNT],
            indents: level_indents(BULLET_INDENT_STEP_POINTS),
            text_indents: vec![BULLET_INDENT_STEP_POINTS / FONT_EM_POINTS; LIST_LEVEL_COUNT],
            ..Default::default()
        };
        let mut encoded = style.encode_to_vec();
        let unknown = [0xb0, 0x06, 0x0b];
        encoded.extend_from_slice(&unknown);
        let color = ParagraphListLabelColor::Explicit(
            crate::shapes::RgbaColor::new(
                0.8,
                0.2,
                0.1,
                0.75,
                crate::shapes::RgbColorSpace::DisplayP3,
            )
            .unwrap(),
        );

        let patched = patch_direct_label_color(&encoded, 10, color).unwrap();
        assert!(
            patched
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        let decoded = tswp::ListStyleArchive::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.override_count, Some(5));
        assert_eq!(decoded.strings, style.strings);
        assert_eq!(decoded.geometries, style.geometries);
        assert_eq!(decoded.indents, style.indents);
        assert_eq!(decoded.text_indents, style.text_indents);
        assert_eq!(
            crate::shapes::color_from_native(decoded.font_color.as_ref().unwrap()).unwrap(),
            match color {
                ParagraphListLabelColor::Explicit(color) => color,
                ParagraphListLabelColor::Automatic => unreachable!(),
            }
        );

        let automatic =
            patch_direct_label_color(&patched, 10, ParagraphListLabelColor::Automatic).unwrap();
        let decoded = tswp::ListStyleArchive::decode(automatic.as_slice()).unwrap();
        assert_eq!(decoded.override_count, Some(5));
        assert_eq!(decoded.font_color_null, Some(true));
        assert!(decoded.font_color.is_none());
    }

    #[test]
    fn every_native_number_format_has_one_strict_public_representation() {
        let mut formats = Vec::new();
        for sequence in ParagraphListNumberSequence::ALL {
            for punctuation in ParagraphListNumberPunctuation::ALL {
                formats.push(ParagraphListNumberFormat::affixed(sequence, punctuation));
            }
        }
        formats.extend([
            ParagraphListNumberFormat::Circled,
            ParagraphListNumberFormat::HebrewBiblicalStandard,
        ]);
        assert_eq!(formats.len(), 65);
        let native = formats
            .iter()
            .copied()
            .map(number_format_to_native)
            .collect::<HashSet<_>>();
        assert_eq!(native.len(), formats.len());
        for format in formats {
            assert_eq!(
                number_format_from_native(number_format_to_native(format)).unwrap(),
                format
            );
        }
        assert!(number_format_from_native(65).is_err());
        assert!(number_format_from_native(-1).is_err());
    }

    #[test]
    fn number_format_updates_compose_losslessly_and_preserve_unknown_fields() {
        let mut style = tswp::ListStyleArchive {
            super_: tss::StyleArchive {
                parent: Some(reference(8)),
                is_variation: Some(true),
                stylesheet: Some(reference(3)),
                ..Default::default()
            },
            override_count: Some(2),
            font_color: Some(crate::shapes::color_to_native(
                crate::shapes::RgbaColor::black(),
            )),
            tiered_numbers: vec![false; LIST_LEVEL_COUNT],
            ..Default::default()
        };
        style.tiered_numbers[1] = true;
        let mut encoded = style.encode_to_vec();
        let unknown = [0xb8, 0x06, 0x0d];
        encoded.extend_from_slice(&unknown);
        let mut formats = repeated_enum(NumberType::KNumericDecimal);
        formats[1] = number_format_to_native(ParagraphListNumberFormat::affixed(
            ParagraphListNumberSequence::RomanLowercase,
            ParagraphListNumberPunctuation::Parentheses,
        ));

        let patched = patch_direct_number_types(&encoded, 10, &formats).unwrap();
        assert!(
            patched
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        let decoded = tswp::ListStyleArchive::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.override_count, Some(3));
        assert_eq!(decoded.number_types, formats);
        assert_eq!(decoded.tiered_numbers, style.tiered_numbers);
        assert_eq!(decoded.font_color, style.font_color);
    }

    #[test]
    fn number_tiering_updates_compose_losslessly_and_preserve_unknown_fields() {
        let formats = repeated_enum(NumberType::KNumericDecimal);
        let style = tswp::ListStyleArchive {
            super_: tss::StyleArchive {
                parent: Some(reference(8)),
                is_variation: Some(true),
                stylesheet: Some(reference(3)),
                ..Default::default()
            },
            override_count: Some(2),
            font_color: Some(crate::shapes::color_to_native(
                crate::shapes::RgbaColor::black(),
            )),
            number_types: formats.clone(),
            ..Default::default()
        };
        let mut encoded = style.encode_to_vec();
        let unknown = [0xb8, 0x06, 0x0d];
        encoded.extend_from_slice(&unknown);
        let mut tiering = vec![false; LIST_LEVEL_COUNT];
        tiering[1] = true;

        let patched = patch_direct_tiered_numbers(&encoded, 10, &tiering).unwrap();
        assert!(
            patched
                .windows(unknown.len())
                .any(|window| window == unknown)
        );
        let decoded = tswp::ListStyleArchive::decode(patched.as_slice()).unwrap();
        assert_eq!(decoded.override_count, Some(3));
        assert_eq!(decoded.tiered_numbers, tiering);
        assert_eq!(decoded.number_types, formats);
        assert_eq!(decoded.font_color, style.font_color);
        assert!(validate_tiered_numbers(10, &[false; LIST_LEVEL_COUNT - 1]).is_err());
    }

    #[test]
    fn native_style_mutations_use_exact_message_anchor_with_sibling_payload() {
        let style_id = 42;
        let mut style = canonical_archive(ParagraphList::Bullet);
        style.super_.parent = Some(reference(8));
        style.super_.is_variation = Some(true);
        let sibling = vec![0x98, 0x06, 0x07];
        let object = ArchiveObject::new(
            style_id,
            vec![
                RawMessage {
                    type_: 2_022,
                    data: sibling.clone(),
                },
                RawMessage {
                    type_: LIST_STYLE_MESSAGE_TYPE,
                    data: style.encode_to_vec(),
                },
            ],
        )
        .unwrap();
        let archive = crate::archive::Archive {
            objects: vec![object],
        };
        let mut package = IWorkPackage::new();
        package.replace_archive("Index/One.iwa", &archive).unwrap();

        let location = locate_style(&package, style_id).unwrap();
        assert_eq!(location.message_index, 1);
        assert_eq!(location.message_type, LIST_STYLE_MESSAGE_TYPE);
        let located = locate_style_with_archive(&package, style_id).unwrap();
        assert!(located.archive.object(style_id).is_some());

        let mut strings = style.strings.clone();
        strings[2] = "◆".to_owned();
        replace_direct_bullet_strings_with_archive(&mut package, located, &strings).unwrap();
        let located_indentation = locate_style_with_archive(&package, style_id).unwrap();
        replace_direct_list_indentation_with_archive(
            &mut package,
            located_indentation,
            &level_indents(BULLET_INDENT_STEP_POINTS),
            &[BULLET_INDENT_STEP_POINTS / FONT_EM_POINTS; LIST_LEVEL_COUNT],
        )
        .unwrap();
        remove_direct_list_indentation(&mut package, &location).unwrap();

        let updated = package.archive("Index/One.iwa").unwrap();
        let updated_object = updated.object(style_id).unwrap();
        assert_eq!(updated_object.messages[0].data, sibling);
        assert_eq!(updated_object.messages[0].type_, 2_022);
        assert_eq!(updated_object.archive_info.message_infos[0].type_, 2_022);
        assert_eq!(
            tswp::ListStyleArchive::decode(updated_object.messages[1].data.as_slice())
                .unwrap()
                .strings,
            strings
        );

        let mut stale = package.clone();
        let stale_located = locate_style_with_archive(&stale, style_id).unwrap();
        stale
            .update_archive("Index/One.iwa", |archive| {
                let object = archive.object_mut(style_id).unwrap();
                object.messages[1].type_ = LIST_STYLE_MESSAGE_TYPE + 1;
                object.archive_info.message_infos[1].type_ = LIST_STYLE_MESSAGE_TYPE + 1;
                Ok(())
            })
            .unwrap();
        let before = stale.entry("Index/One.iwa").unwrap().to_vec();
        assert!(
            replace_direct_bullet_strings_with_archive(&mut stale, stale_located, &strings)
                .is_err()
        );
        assert_eq!(stale.entry("Index/One.iwa").unwrap(), before.as_slice());
    }
}
