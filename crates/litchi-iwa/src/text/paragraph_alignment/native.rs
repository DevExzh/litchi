//! Native paragraph-style inheritance, minimal variations, and ownership checks.

use std::collections::HashSet;

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::{tsp, tss, tswp};
use crate::wire::{parse_wire_fields, repeated_length_delimited_payloads};
use crate::{Error, IWorkPackage, Result};

use super::super::style::{
    ParagraphLineSpacing, ParagraphLineSpacingMultiple, ParagraphLineSpacingPoints,
    ParagraphSpacing, ParagraphSpacingPoints, TextAlignment,
};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_CHARACTER_PROPERTIES_FIELD: u32 = 11;
const STYLE_PARAGRAPH_PROPERTIES_FIELD: u32 = 12;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const PARAGRAPH_ALIGNMENT_FIELD: u32 = 1;
const PARAGRAPH_LINE_SPACING_FIELD: u32 = 13;
const PARAGRAPH_SPACE_AFTER_FIELD: u32 = 20;
const PARAGRAPH_SPACE_BEFORE_FIELD: u32 = 21;
const LINE_SPACING_MODE_FIELD: u32 = 1;
const LINE_SPACING_AMOUNT_FIELD: u32 = 2;

const RELATIVE_LINE_SPACING_MODE: i32 = 0;
const MINIMUM_LINE_SPACING_MODE: i32 = 1;
const EXACT_LINE_SPACING_MODE: i32 = 2;
const MAXIMUM_LINE_SPACING_MODE: i32 = 3;
const BETWEEN_LINE_SPACING_MODE: i32 = 4;

pub(super) struct ParagraphStyleLocation {
    pub(super) archive_name: String,
    pub(super) message: RawMessage,
    pub(super) style: tswp::ParagraphStyleArchive,
}

#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub(super) struct ParagraphStyleOverrides {
    pub(super) alignment: Option<TextAlignment>,
    pub(super) line_spacing: Option<ParagraphLineSpacing>,
    pub(super) space_before: Option<ParagraphSpacingPoints>,
    pub(super) space_after: Option<ParagraphSpacingPoints>,
}

impl ParagraphStyleOverrides {
    pub(super) fn count(self) -> u32 {
        u32::from(self.alignment.is_some())
            + u32::from(self.line_spacing.is_some())
            + u32::from(self.space_before.is_some())
            + u32::from(self.space_after.is_some())
    }

    pub(super) fn is_empty(self) -> bool {
        self.count() == 0
    }
}

pub(super) fn locate_style(
    package: &IWorkPackage,
    style_id: u64,
) -> Result<ParagraphStyleLocation> {
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
    Ok(ParagraphStyleLocation {
        archive_name,
        message: message.clone(),
        style: style.clone(),
    })
}

pub(super) fn inherited_alignment(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<TextAlignment> {
    walk_inheritance(package, first_style_id, |style| {
        style
            .para_properties
            .as_ref()
            .and_then(|properties| properties.alignment)
            .map(TextAlignment::from_native_value)
            .transpose()
    })?
    .map_or(Ok(TextAlignment::Natural), Ok)
}

pub(super) fn inherited_line_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphLineSpacing> {
    walk_inheritance(package, first_style_id, |style| {
        let Some(properties) = style.para_properties.as_ref() else {
            return Ok(None);
        };
        if properties.line_spacing_null == Some(true) {
            return Ok(Some(ParagraphLineSpacing::default()));
        }
        properties
            .line_spacing
            .as_ref()
            .map(line_spacing_from_archive)
            .transpose()
    })?
    .map_or(Ok(ParagraphLineSpacing::default()), Ok)
}

pub(super) fn inherited_spacing(
    package: &IWorkPackage,
    first_style_id: u64,
) -> Result<ParagraphSpacing> {
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    let mut before = None;
    let mut after = None;
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(ParagraphSpacing::new(
                before.unwrap_or(ParagraphSpacingPoints::ZERO),
                after.unwrap_or(ParagraphSpacingPoints::ZERO),
            ));
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style inheritance cycles at {identifier}"
            )));
        }
        let location = locate_style(package, identifier)?;
        if let Some(properties) = location.style.para_properties.as_ref() {
            if before.is_none() {
                before = properties
                    .space_before
                    .map(ParagraphSpacingPoints::from_points)
                    .transpose()?;
            }
            if after.is_none() {
                after = properties
                    .space_after
                    .map(ParagraphSpacingPoints::from_points)
                    .transpose()?;
            }
            if let (Some(before), Some(after)) = (before, after) {
                return Ok(ParagraphSpacing::new(before, after));
            }
        }
        style_id = location.style.super_.parent.map(|parent| parent.identifier);
    }
    Err(Error::InvalidFormat(format!(
        "iWork paragraph style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}

fn walk_inheritance<T, F>(
    package: &IWorkPackage,
    first_style_id: u64,
    mut property: F,
) -> Result<Option<T>>
where
    F: FnMut(&tswp::ParagraphStyleArchive) -> Result<Option<T>>,
{
    let mut visited = HashSet::new();
    let mut style_id = Some(first_style_id);
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        let Some(identifier) = style_id else {
            return Ok(None);
        };
        if !visited.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork paragraph style inheritance cycles at {identifier}"
            )));
        }
        let location = locate_style(package, identifier)?;
        if let Some(value) = property(&location.style)? {
            return Ok(Some(value));
        }
        style_id = location.style.super_.parent.map(|parent| parent.identifier);
    }
    Err(Error::InvalidFormat(format!(
        "iWork paragraph style inheritance exceeds {MAX_STYLE_INHERITANCE_DEPTH} levels"
    )))
}

pub(super) fn direct_overrides(
    style: &tswp::ParagraphStyleArchive,
    raw: &[u8],
) -> Result<Option<ParagraphStyleOverrides>> {
    let Some(properties) = style.para_properties.as_ref() else {
        return Ok(None);
    };
    let alignment = properties
        .alignment
        .map(TextAlignment::from_native_value)
        .transpose()?;
    let line_spacing = properties
        .line_spacing
        .as_ref()
        .map(line_spacing_from_archive)
        .transpose()?;
    let space_before = properties
        .space_before
        .map(ParagraphSpacingPoints::from_points)
        .transpose()?;
    let space_after = properties
        .space_after
        .map(ParagraphSpacingPoints::from_points)
        .transpose()?;
    let overrides = ParagraphStyleOverrides {
        alignment,
        line_spacing,
        space_before,
        space_after,
    };
    let mut remaining = properties.clone();
    remaining.alignment = None;
    remaining.line_spacing = None;
    remaining.space_before = None;
    remaining.space_after = None;
    let semantic = !overrides.is_empty()
        && remaining == tswp::ParagraphStylePropertiesArchive::default()
        && style.override_count == Some(overrides.count())
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
        return Ok(None);
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
    let mut paragraph_fields = Vec::with_capacity(4);
    if alignment.is_some() {
        paragraph_fields.push(PARAGRAPH_ALIGNMENT_FIELD);
    }
    if line_spacing.is_some() {
        paragraph_fields.push(PARAGRAPH_LINE_SPACING_FIELD);
        let line_spacing_raw = required_payload(
            paragraph_raw,
            PARAGRAPH_LINE_SPACING_FIELD,
            "paragraph line spacing",
        )?;
        let expected = match line_spacing {
            Some(ParagraphLineSpacing::Relative(_)) => vec![LINE_SPACING_AMOUNT_FIELD],
            Some(_) => vec![LINE_SPACING_MODE_FIELD, LINE_SPACING_AMOUNT_FIELD],
            None => Vec::new(),
        };
        if !has_exact_fields(line_spacing_raw, &expected)? {
            return Ok(None);
        }
    }
    if space_after.is_some() {
        paragraph_fields.push(PARAGRAPH_SPACE_AFTER_FIELD);
    }
    if space_before.is_some() {
        paragraph_fields.push(PARAGRAPH_SPACE_BEFORE_FIELD);
    }
    let exact = has_exact_fields(
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
        && has_exact_fields(paragraph_raw, &paragraph_fields)?;
    Ok(exact.then_some(overrides))
}

pub(super) fn variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    overrides: ParagraphStyleOverrides,
) -> Result<ArchiveObject> {
    if overrides.is_empty() {
        return Err(Error::InvalidFormat(
            "an iWork paragraph-style variation must contain an override".to_owned(),
        ));
    }
    let data = tswp::ParagraphStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(overrides.count()),
        char_properties: Some(tswp::CharacterStylePropertiesArchive::default()),
        para_properties: Some(tswp::ParagraphStylePropertiesArchive {
            alignment: overrides.alignment.map(TextAlignment::native_value),
            line_spacing: overrides.line_spacing.map(line_spacing_archive),
            space_before: overrides.space_before.map(ParagraphSpacingPoints::points),
            space_after: overrides.space_after.map(ParagraphSpacingPoints::points),
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

pub(super) fn replace_variation(
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

pub(super) fn is_exclusive(package: &IWorkPackage, style_id: u64) -> Result<bool> {
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

pub(super) fn parent_style_id(style: &tswp::ParagraphStyleArchive, style_id: u64) -> Result<u64> {
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

pub(super) fn stylesheet_id(style: &tswp::ParagraphStyleArchive, style_id: u64) -> Result<u64> {
    style
        .super_
        .stylesheet
        .as_ref()
        .map(|stylesheet| stylesheet.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {style_id} has no stylesheet"
            ))
        })
}

pub(super) fn object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
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

fn line_spacing_archive(spacing: ParagraphLineSpacing) -> tswp::LineSpacingArchive {
    match spacing {
        ParagraphLineSpacing::Relative(multiple) => tswp::LineSpacingArchive {
            amount: Some(multiple.get()),
            ..Default::default()
        },
        ParagraphLineSpacing::AtLeast(points) => point_spacing(MINIMUM_LINE_SPACING_MODE, points),
        ParagraphLineSpacing::Exactly(points) => point_spacing(EXACT_LINE_SPACING_MODE, points),
        ParagraphLineSpacing::Maximum(points) => point_spacing(MAXIMUM_LINE_SPACING_MODE, points),
        ParagraphLineSpacing::Between(points) => point_spacing(BETWEEN_LINE_SPACING_MODE, points),
    }
}

fn point_spacing(mode: i32, points: ParagraphLineSpacingPoints) -> tswp::LineSpacingArchive {
    tswp::LineSpacingArchive {
        mode: Some(mode),
        amount: Some(points.points()),
        ..Default::default()
    }
}

fn line_spacing_from_archive(spacing: &tswp::LineSpacingArchive) -> Result<ParagraphLineSpacing> {
    if spacing.baseline_rule.is_some() {
        return Err(Error::InvalidFormat(
            "unsupported native iWork line-spacing baseline rule".to_owned(),
        ));
    }
    let mode = spacing.mode.unwrap_or(RELATIVE_LINE_SPACING_MODE);
    if mode == RELATIVE_LINE_SPACING_MODE {
        let multiple = spacing.amount.unwrap_or(1.0);
        return Ok(ParagraphLineSpacing::Relative(
            ParagraphLineSpacingMultiple::new(multiple)?,
        ));
    }
    let points = spacing.amount.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "native iWork line-spacing mode {mode} has no amount"
        ))
    })?;
    let points = ParagraphLineSpacingPoints::from_points(points)?;
    match mode {
        MINIMUM_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::AtLeast(points)),
        EXACT_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::Exactly(points)),
        MAXIMUM_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::Maximum(points)),
        BETWEEN_LINE_SPACING_MODE => Ok(ParagraphLineSpacing::Between(points)),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported native iWork line-spacing mode {mode}"
        ))),
    }
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

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
