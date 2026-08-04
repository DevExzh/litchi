//! Canonical native Drop Cap style variations and ownership checks.

use prost::Message;

use crate::archive::{ArchiveObject, RawMessage};
use crate::protobuf::{tsp, tss, tswp};
use crate::wire::{parse_wire_fields, repeated_length_delimited_payloads};
use crate::{Error, IWorkPackage, IWorkThemeArchive, Result};

use super::types::{
    DropCapCharacterCount, DropCapCharacterScale, DropCapCornerRadius, DropCapLineCount,
    DropCapOutdent, DropCapPadding, DropCapRaisedLines, DropCapWrap, ParagraphDropCap,
};
use crate::text::style_registry::{object_archive, object_archive_name};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const DROP_CAP_STYLE_MESSAGE_TYPE: u32 = 10_024;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_CHARACTER_PROPERTIES_FIELD: u32 = 11;
const STYLE_DROP_CAP_PROPERTIES_FIELD: u32 = 12;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const DROP_CAP_MODEL_FIELD: u32 = 1;

const MODEL_TYPE_FIELD: u32 = 1;
const MODEL_LINES_FIELD: u32 = 2;
const MODEL_RAISED_LINES_FIELD: u32 = 3;
const MODEL_WRAP_FIELD: u32 = 6;
const MODEL_SHAPE_ENABLED_FIELD: u32 = 7;
const MODEL_CHARACTERS_FIELD: u32 = 10;
const MODEL_OUTDENT_FIELD: u32 = 11;
const MODEL_PADDING_FIELD: u32 = 12;
const MODEL_CORNER_RADIUS_FIELD: u32 = 13;
const MODEL_CHARACTER_SCALE_FIELD: u32 = 14;

const TEXT_DROP_CAP_TYPE: i32 = 0;

pub(super) struct DropCapStyleLocation {
    pub(super) object_id: u64,
    pub(super) archive_name: String,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) message: RawMessage,
    pub(super) style: tswp::DropCapStyleArchive,
}

pub(super) struct DropCapBaseStyle {
    pub(super) archive_name: String,
    pub(super) identifier: u64,
}

pub(super) fn locate_style(package: &IWorkPackage, style_id: u64) -> Result<DropCapStyleLocation> {
    let (archive_name, archive) = object_archive(package, style_id)?;
    let object = archive.object(style_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork Drop Cap style {style_id} is missing"))
    })?;
    let payloads = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == DROP_CAP_STYLE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [(message_index, message)] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork Drop Cap style {style_id} must have exactly one DropCapStyle payload"
        )));
    };
    let style = tswp::DropCapStyleArchive::decode(message.data.as_slice())?;
    Ok(DropCapStyleLocation {
        object_id: style_id,
        archive_name,
        message_index: *message_index,
        message_type: message.type_,
        message: (*message).clone(),
        style,
    })
}

pub(super) fn plain_text_model(
    style_id: u64,
    location: &DropCapStyleLocation,
) -> Result<ParagraphDropCap> {
    let style = &location.style;
    let properties = style.drop_cap_properties.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork Drop Cap style {style_id} has no Drop Cap properties"
        ))
    })?;
    let model = properties.drop_cap.as_ref().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork Drop Cap style {style_id} has no Drop Cap model"
        ))
    })?;
    let drop_cap = model_from_archive(model)?;
    if style.override_count != Some(1)
        || style.char_properties.as_ref() != Some(&tswp::CharacterStylePropertiesArchive::default())
        || properties.drop_cap_shape_stroke.is_some()
        || properties.drop_cap_shape_fill_null.is_some()
        || properties.drop_cap_shape_fill.is_some()
        || style.super_.name.is_some()
        || style.super_.style_identifier.is_some()
        || style.super_.parent.is_none()
        || style.super_.is_variation != Some(true)
        || style.super_.stylesheet.is_none()
        || !has_canonical_wire(&location.message.data)?
    {
        return Err(Error::InvalidFormat(format!(
            "iWork Drop Cap style {style_id} is not a canonical plain-text variation"
        )));
    }
    Ok(drop_cap)
}

pub(super) fn base_style(package: &IWorkPackage, stylesheet_id: u64) -> Result<DropCapBaseStyle> {
    let archive_name = object_archive_name(package, stylesheet_id)?;
    let mut identifiers = Vec::new();
    for candidate_archive_name in package.iwa_entry_names() {
        for object in package.archive(candidate_archive_name)?.objects {
            for message in &object.messages {
                let Ok(theme) = IWorkThemeArchive::decode(&message.data) else {
                    continue;
                };
                identifiers.extend(
                    theme
                        .extensions
                        .text
                        .into_iter()
                        .flat_map(|presets| presets.dropcap_style_presets)
                        .map(|reference| reference.identifier)
                        .filter(|identifier| *identifier != 0),
                );
            }
        }
    }
    identifiers.sort_unstable();
    identifiers.dedup();
    identifiers.retain(|identifier| {
        locate_style(package, *identifier).is_ok_and(|base| {
            base.archive_name == archive_name
                && base.style.super_.parent.is_none()
                && base.style.super_.is_variation.is_none()
                && base
                    .style
                    .super_
                    .stylesheet
                    .as_ref()
                    .map(|reference| reference.identifier)
                    == Some(stylesheet_id)
        })
    });
    let [identifier] = identifiers.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork stylesheet {stylesheet_id} must identify one Drop Cap base style"
        )));
    };
    Ok(DropCapBaseStyle {
        archive_name,
        identifier: *identifier,
    })
}

pub(super) fn variation_object(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    drop_cap: ParagraphDropCap,
) -> Result<ArchiveObject> {
    let data = tswp::DropCapStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        char_properties: Some(tswp::CharacterStylePropertiesArchive::default()),
        drop_cap_properties: Some(tswp::DropCapStylePropertiesArchive {
            drop_cap: Some(model_archive(drop_cap)),
            ..Default::default()
        }),
    }
    .encode_to_vec();
    let decoded = tswp::DropCapStyleArchive::decode(data.as_slice())?;
    let location = DropCapStyleLocation {
        object_id: identifier,
        archive_name: String::new(),
        message_index: 0,
        message_type: DROP_CAP_STYLE_MESSAGE_TYPE,
        message: RawMessage {
            type_: DROP_CAP_STYLE_MESSAGE_TYPE,
            data: data.clone(),
        },
        style: decoded,
    };
    if plain_text_model(identifier, &location)? != drop_cap {
        return Err(Error::InvalidFormat(
            "generated iWork Drop Cap variation failed canonical validation".to_owned(),
        ));
    }
    let mut object = ArchiveObject::new(identifier, vec![location.message])?;
    object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    object.archive_info.message_infos[0]
        .object_references
        .push(parent_style_id);
    Ok(object)
}

pub(super) fn replace_variation(
    package: &mut IWorkPackage,
    location: &DropCapStyleLocation,
    mut replacement: ArchiveObject,
) -> Result<()> {
    let style_id = location.object_id;
    if replacement.archive_info.identifier != Some(style_id)
        || replacement.messages.len() != 1
        || replacement.archive_info.message_infos.len() != 1
    {
        return Err(Error::InvalidFormat(
            "replacement Drop Cap style does not contain exactly one object-aligned payload"
                .to_owned(),
        ));
    }
    if replacement.messages[0].type_ != location.message_type {
        return Err(Error::InvalidFormat(
            "replacement Drop Cap style payload type does not match its anchor".to_owned(),
        ));
    }
    if replacement.archive_info.message_infos[0].type_ != location.message_type {
        return Err(Error::InvalidFormat(
            "replacement Drop Cap style metadata type does not match its anchor".to_owned(),
        ));
    }
    let message = replacement.messages.pop().ok_or_else(|| {
        Error::InvalidFormat("replacement Drop Cap style has no payload".to_owned())
    })?;
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork Drop Cap style {style_id} is missing"))
        })?;
        if object.archive_info.identifier != Some(style_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork Drop Cap style {style_id} object identity changed unexpectedly"
            )));
        }
        if object.messages.get(location.message_index).is_none() {
            return Err(Error::InvalidFormat(format!(
                "iWork Drop Cap style {style_id} anchored payload index {} is missing",
                location.message_index
            )));
        }
        if object.messages[location.message_index].type_ != location.message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork Drop Cap style {style_id} anchored payload type changed unexpectedly"
            )));
        }
        if object.messages[location.message_index].data != location.message.data {
            return Err(Error::InvalidFormat(format!(
                "iWork Drop Cap style {style_id} anchored payload changed unexpectedly"
            )));
        }
        let Some(info) = object
            .archive_info
            .message_infos
            .get(location.message_index)
        else {
            return Err(Error::InvalidFormat(format!(
                "iWork Drop Cap style {style_id} anchored metadata index {} is missing",
                location.message_index
            )));
        };
        if info.type_ != location.message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork Drop Cap style {style_id} anchored metadata type changed unexpectedly"
            )));
        }
        object.replace_message(location.message_index, message)?;
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
                        .table_drop_cap_style
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
                if message.type_ == DROP_CAP_STYLE_MESSAGE_TYPE
                    && let Ok(style) = tswp::DropCapStyleArchive::decode(message.data.as_slice())
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

pub(super) fn parent_style_id(style: &tswp::DropCapStyleArchive, style_id: u64) -> Result<u64> {
    style
        .super_
        .parent
        .as_ref()
        .map(|parent| parent.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork Drop Cap variation {style_id} has no parent"))
        })
}

pub(super) fn stylesheet_id(style: &tswp::DropCapStyleArchive, style_id: u64) -> Result<u64> {
    style
        .super_
        .stylesheet
        .as_ref()
        .map(|stylesheet| stylesheet.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork Drop Cap style {style_id} has no stylesheet"))
        })
}

fn model_from_archive(model: &tswp::DropCapArchive) -> Result<ParagraphDropCap> {
    if model.r#type != Some(TEXT_DROP_CAP_TYPE)
        || model.shape_enabled != Some(false)
        || model.deprecated_outdent.is_some()
        || model.deprecated_padding.is_some()
        || model.deprecated_corner_radius.is_some()
        || model.deprecated_character_scale.is_some()
    {
        return Err(Error::InvalidFormat(
            "unsupported native iWork Drop Cap model".to_owned(),
        ));
    }
    let lines = u8::try_from(required(model.number_of_lines, "line count")?).map_err(|_| {
        Error::InvalidFormat("native iWork Drop Cap line count exceeds u8".to_owned())
    })?;
    let raised_lines = u8::try_from(required(model.number_of_raised_lines, "raised-line count")?)
        .map_err(|_| {
        Error::InvalidFormat("native iWork Drop Cap raised-line count exceeds u8".to_owned())
    })?;
    Ok(ParagraphDropCap {
        lines: DropCapLineCount::new(lines)?,
        characters: DropCapCharacterCount::new(required(
            model.number_of_characters,
            "character count",
        )?)?,
        raised_lines: DropCapRaisedLines::new(raised_lines)?,
        wrap: DropCapWrap::from_native_value(required(model.wrap_type, "wrap type")?)?,
        padding: DropCapPadding::from_points(required(model.padding, "padding")?)?,
        outdent: DropCapOutdent::from_ratio(required(model.outdent, "outdent")?)?,
        corner_radius: DropCapCornerRadius::from_ratio(required(
            model.corner_radius,
            "corner radius",
        )?)?,
        character_scale: DropCapCharacterScale::from_ratio(required(
            model.character_scale,
            "character scale",
        )?)?,
    })
}

fn model_archive(drop_cap: ParagraphDropCap) -> tswp::DropCapArchive {
    tswp::DropCapArchive {
        r#type: Some(TEXT_DROP_CAP_TYPE),
        number_of_lines: Some(u32::from(drop_cap.lines.get())),
        number_of_raised_lines: Some(u32::from(drop_cap.raised_lines.get())),
        outdent: Some(drop_cap.outdent.ratio()),
        padding: Some(drop_cap.padding.points()),
        wrap_type: Some(drop_cap.wrap.native_value()),
        shape_enabled: Some(false),
        corner_radius: Some(drop_cap.corner_radius.ratio()),
        character_scale: Some(drop_cap.character_scale.ratio()),
        number_of_characters: Some(drop_cap.characters.get()),
        ..Default::default()
    }
}

fn has_canonical_wire(data: &[u8]) -> Result<bool> {
    if !has_exact_fields(
        data,
        &[
            STYLE_SUPER_FIELD,
            STYLE_OVERRIDE_COUNT_FIELD,
            STYLE_CHARACTER_PROPERTIES_FIELD,
            STYLE_DROP_CAP_PROPERTIES_FIELD,
        ],
    )? {
        return Ok(false);
    }
    let super_raw = required_payload(data, STYLE_SUPER_FIELD, "Drop Cap style")?;
    let character_raw = required_payload(
        data,
        STYLE_CHARACTER_PROPERTIES_FIELD,
        "Drop Cap character properties",
    )?;
    let properties_raw =
        required_payload(data, STYLE_DROP_CAP_PROPERTIES_FIELD, "Drop Cap properties")?;
    let model_raw = required_payload(properties_raw, DROP_CAP_MODEL_FIELD, "Drop Cap model")?;
    Ok(has_exact_fields(
        super_raw,
        &[
            STYLE_PARENT_FIELD,
            STYLE_VARIATION_FIELD,
            STYLE_STYLESHEET_FIELD,
        ],
    )? && has_exact_fields(character_raw, &[])?
        && has_exact_fields(properties_raw, &[DROP_CAP_MODEL_FIELD])?
        && has_exact_fields(
            model_raw,
            &[
                MODEL_TYPE_FIELD,
                MODEL_LINES_FIELD,
                MODEL_RAISED_LINES_FIELD,
                MODEL_WRAP_FIELD,
                MODEL_SHAPE_ENABLED_FIELD,
                MODEL_CHARACTERS_FIELD,
                MODEL_OUTDENT_FIELD,
                MODEL_PADDING_FIELD,
                MODEL_CORNER_RADIUS_FIELD,
                MODEL_CHARACTER_SCALE_FIELD,
            ],
        )?)
}

fn has_exact_fields(data: &[u8], expected: &[u32]) -> Result<bool> {
    let mut actual = parse_wire_fields(data)?
        .into_iter()
        .map(|field| field.number())
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

fn required<T>(value: Option<T>, label: &str) -> Result<T> {
    value.ok_or_else(|| Error::InvalidFormat(format!("native iWork Drop Cap has no {label}")))
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
    fn canonical_plain_text_model_round_trips() {
        let expected = ParagraphDropCap::new(
            DropCapLineCount::new(4).unwrap(),
            DropCapCharacterCount::new(2).unwrap(),
        )
        .with_raised_lines(DropCapRaisedLines::new(1).unwrap())
        .with_wrap(DropCapWrap::Contour)
        .with_padding(DropCapPadding::from_points(6.0).unwrap())
        .with_outdent(DropCapOutdent::from_ratio(0.25).unwrap());
        let object = variation_object(9, 7, 3, expected).unwrap();
        let message = object.messages[0].clone();
        let location = DropCapStyleLocation {
            object_id: 9,
            archive_name: String::new(),
            message_index: 0,
            message_type: DROP_CAP_STYLE_MESSAGE_TYPE,
            style: tswp::DropCapStyleArchive::decode(message.data.as_slice()).unwrap(),
            message,
        };
        assert_eq!(plain_text_model(9, &location).unwrap(), expected);
    }

    #[test]
    fn style_replacement_uses_exact_message_anchor_with_sibling_payload() {
        let original_model = ParagraphDropCap::new(
            DropCapLineCount::new(3).unwrap(),
            DropCapCharacterCount::new(1).unwrap(),
        );
        let replacement_model = ParagraphDropCap::new(
            DropCapLineCount::new(5).unwrap(),
            DropCapCharacterCount::new(2).unwrap(),
        );
        let original = variation_object(9, 7, 3, original_model).unwrap();
        let original_message = original.messages[0].clone();
        let sibling = tswp::ParagraphStyleArchive::default().encode_to_vec();
        let object = ArchiveObject::new(
            9,
            vec![
                RawMessage {
                    type_: 2_022,
                    data: sibling.clone(),
                },
                original_message,
            ],
        )
        .unwrap();
        let archive = crate::archive::Archive {
            objects: vec![object],
        };
        let mut package = IWorkPackage::new();
        package.replace_archive("Index/One.iwa", &archive).unwrap();

        let location = locate_style(&package, 9).unwrap();
        assert_eq!(location.message_index, 1);
        assert_eq!(location.message_type, DROP_CAP_STYLE_MESSAGE_TYPE);
        replace_variation(
            &mut package,
            &location,
            variation_object(9, 7, 3, replacement_model).unwrap(),
        )
        .unwrap();

        let updated = package.archive("Index/One.iwa").unwrap();
        let updated_object = updated.object(9).unwrap();
        assert_eq!(updated_object.messages[0].data, sibling);
        assert_eq!(updated_object.messages[0].type_, 2_022);
        assert_eq!(
            plain_text_model(9, &locate_style(&package, 9).unwrap()).unwrap(),
            replacement_model
        );

        let mut stale = package.clone();
        stale
            .update_archive("Index/One.iwa", |archive| {
                let object = archive.object_mut(9).unwrap();
                object.messages[1].type_ = DROP_CAP_STYLE_MESSAGE_TYPE + 1;
                object.archive_info.message_infos[1].type_ = DROP_CAP_STYLE_MESSAGE_TYPE + 1;
                Ok(())
            })
            .unwrap();
        let before = stale.entry("Index/One.iwa").unwrap().to_vec();
        assert!(
            replace_variation(
                &mut stale,
                &location,
                variation_object(9, 7, 3, original_model).unwrap(),
            )
            .is_err()
        );
        assert_eq!(stale.entry("Index/One.iwa").unwrap(), before.as_slice());
    }

    #[test]
    fn shaped_and_noncanonical_models_are_rejected() {
        let mut shaped = model_archive(ParagraphDropCap::default());
        shaped.shape_enabled = Some(true);
        assert!(model_from_archive(&shaped).is_err());

        let mut missing_scale = model_archive(ParagraphDropCap::default());
        missing_scale.character_scale = None;
        assert!(model_from_archive(&missing_scale).is_err());
    }
}
