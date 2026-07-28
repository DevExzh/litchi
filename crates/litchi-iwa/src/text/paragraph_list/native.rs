//! Canonical native TSWP list-style objects.

use std::borrow::Cow;
use std::collections::HashSet;

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::tswp::list_style_archive::{LabelGeometry, LabelType, NumberType};
use crate::protobuf::{tsp, tss, tswp};
use crate::{Error, IWorkPackage, Result};

use super::types::ParagraphList;

const LIST_STYLE_MESSAGE_TYPE: u32 = 2_023;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const LIST_LEVEL_COUNT: usize = 9;
const FONT_EM_POINTS: f32 = 11.0;
const NONE_INDENT_STEP_POINTS: f32 = 36.0;
const BULLET_INDENT_STEP_POINTS: f32 = 9.0;
const NUMBER_INDENT_STEP_POINTS: f32 = 18.0;
const BULLET_BASELINE_OFFSET_POINTS: f32 = -1.0;
const DEFAULT_LABEL_SCALE: f32 = 1.0;
const BULLET_GLYPH: &str = "•";
const NONE_OVERRIDE_COUNT: u32 = 4;
const BULLET_OVERRIDE_COUNT: u32 = 5;
const NUMBER_OVERRIDE_COUNT: u32 = 6;

pub(super) struct ListStyleLocation {
    pub(super) archive_name: String,
    pub(super) style: tswp::ListStyleArchive,
}

pub(super) fn locate_style(package: &IWorkPackage, style_id: u64) -> Result<ListStyleLocation> {
    let mut found = None;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(style_id) else {
            continue;
        };
        let mut styles = object
            .messages
            .iter()
            .filter(|message| message.type_ == LIST_STYLE_MESSAGE_TYPE)
            .map(|message| tswp::ListStyleArchive::decode(message.data.as_slice()))
            .collect::<std::result::Result<Vec<_>, _>>()?;
        if styles.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork list style {style_id} must have exactly one writable payload"
            )));
        }
        let style = styles.pop().ok_or_else(|| {
            Error::InvalidFormat(format!("iWork list style {style_id} payload disappeared"))
        })?;
        if found
            .replace(ListStyleLocation {
                archive_name: archive_name.to_owned(),
                style,
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

pub(super) fn find_preset_style(
    package: &IWorkPackage,
    archive_name: &str,
    stylesheet_id: u64,
    preset: ParagraphList,
) -> Result<Option<u64>> {
    let archive = package.archive(archive_name)?;
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
}
