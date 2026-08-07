//! Focused wire patches for legend paragraph-style selection and typography.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::ChartFontSize;
use crate::charts::legend_style::{
    GENERATED_LEGEND_STYLE_EXTENSION_FIELD, generated_legend_style_extension,
};
use crate::protobuf::tsch;
use crate::text::TextFont;
use crate::text::paragraph_alignment::native::{direct_overrides, locate_style};
use crate::wire::{
    patch_fixed32_field, patch_length_delimited_field, patch_varint_field,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

/// `tschlegendmodeldefaultlabelparagraphstyleindex`.
const LEGEND_LABEL_PARAGRAPH_STYLE_INDEX_FIELD: u32 = 2;
/// `TSWP.ParagraphStyleArchive.override_count`.
const PARAGRAPH_OVERRIDE_COUNT_FIELD: u32 = 10;
/// `TSWP.ParagraphStyleArchive.char_properties`.
const PARAGRAPH_CHARACTER_PROPERTIES_FIELD: u32 = 11;
/// `TSWP.CharacterStylePropertiesArchive.font_size`.
const CHARACTER_FONT_SIZE_FIELD: u32 = 3;
/// `TSWP.CharacterStylePropertiesArchive.bold`.
const CHARACTER_BOLD_FIELD: u32 = 1;
/// `TSWP.CharacterStylePropertiesArchive.italic`.
const CHARACTER_ITALIC_FIELD: u32 = 2;
/// `TSWP.CharacterStylePropertiesArchive.font_name_null`.
const CHARACTER_FONT_NAME_NULL_FIELD: u32 = 4;
/// `TSWP.CharacterStylePropertiesArchive.font_name`.
const CHARACTER_FONT_NAME_FIELD: u32 = 5;
const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;

pub(super) fn direct_paragraph_style_index(data: &[u8]) -> Result<Option<usize>> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        return Ok(None);
    };
    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    generated
        .tschlegendmodeldefaultlabelparagraphstyleindex
        .map(|index| {
            usize::try_from(index).map_err(|_| {
                Error::InvalidFormat(format!(
                    "native chart legend paragraph-style index {index} is negative"
                ))
            })
        })
        .transpose()
}

pub(super) fn patch_direct_paragraph_style_index(
    data: &[u8],
    index: Option<u64>,
) -> Result<Vec<u8>> {
    let Some(extension) = generated_legend_style_extension(data)? else {
        let Some(index) = index else {
            return Ok(data.to_vec());
        };
        let native_index = i32::try_from(index).map_err(|_| {
            Error::InvalidFormat("legend paragraph-style index exceeds i32".to_owned())
        })?;
        let generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultlabelparagraphstyleindex: Some(native_index),
            ..Default::default()
        }
        .encode_to_vec();
        return patch_length_delimited_field(
            data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            false,
            Some(&generated),
        );
    };
    let generated = tsch::generated::LegendStyleArchive::decode(extension)?;
    let present = generated
        .tschlegendmodeldefaultlabelparagraphstyleindex
        .is_some();
    let patched_extension = patch_varint_field(
        extension,
        LEGEND_LABEL_PARAGRAPH_STYLE_INDEX_FIELD,
        present,
        index,
    )?;
    patch_length_delimited_field(
        data,
        GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
        true,
        Some(&patched_extension),
    )
}

pub(super) fn patch_existing_size(
    package: &mut IWorkPackage,
    style_id: u64,
    size: Option<ChartFontSize>,
) -> Result<()> {
    let location = locate_style(package, style_id)?;
    let overrides =
        direct_overrides(&location.style, &location.message.data)?.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "legend paragraph style {style_id} is not an exact native variation"
            ))
        })?;
    let was_present = overrides.point_size.is_some();
    let mut expected = overrides;
    expected.point_size = size.map(ChartFontSize::text_point_size);
    let next_count = u64::from(expected.count());
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("legend paragraph style {style_id} is missing"))
        })?;
        let messages = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        let [(message_index, message)] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "legend paragraph style {style_id} must have exactly one paragraph-style payload"
            )));
        };
        let data = transform_length_delimited_field(
            &message.data,
            PARAGRAPH_CHARACTER_PROPERTIES_FIELD,
            |characters| {
                patch_fixed32_field(
                    characters,
                    CHARACTER_FONT_SIZE_FIELD,
                    was_present,
                    size.map(|value| value.points().to_bits()),
                )
            },
        )?;
        let data = patch_varint_field(
            &data,
            PARAGRAPH_OVERRIDE_COUNT_FIELD,
            location.style.override_count.is_some(),
            Some(next_count),
        )?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;
    let verified = locate_style(package, style_id)?;
    let actual = direct_overrides(&verified.style, &verified.message.data)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "legend paragraph style {style_id} stopped being an exact native variation"
        ))
    })?;
    if actual != expected {
        return Err(Error::InvalidFormat(format!(
            "legend paragraph style {style_id} font-size wire patch failed validation"
        )));
    }
    Ok(())
}

pub(super) fn patch_existing_font(
    package: &mut IWorkPackage,
    style_id: u64,
    expected: &crate::text::paragraph_alignment::native::ParagraphStyleOverrides,
) -> Result<()> {
    let location = locate_style(package, style_id)?;
    let current = direct_overrides(&location.style, &location.message.data)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "legend paragraph style {style_id} is not an exact native variation"
        ))
    })?;
    let current_font_null = matches!(current.font, Some(TextFont::Default));
    let current_font_name = matches!(current.font, Some(TextFont::Named(_)));
    let expected_font_null = matches!(expected.font, Some(TextFont::Default));
    let expected_font_name = expected.font.as_ref().and_then(TextFont::name);
    let next_count = u64::from(expected.count());
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("legend paragraph style {style_id} is missing"))
        })?;
        let messages = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        let [(message_index, message)] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "legend paragraph style {style_id} must have exactly one paragraph-style payload"
            )));
        };
        let data = transform_length_delimited_field(
            &message.data,
            PARAGRAPH_CHARACTER_PROPERTIES_FIELD,
            |characters| {
                let characters = patch_varint_field(
                    characters,
                    CHARACTER_BOLD_FIELD,
                    current.bold.is_some(),
                    expected.bold.map(u64::from),
                )?;
                let characters = patch_varint_field(
                    &characters,
                    CHARACTER_ITALIC_FIELD,
                    current.italic.is_some(),
                    expected.italic.map(u64::from),
                )?;
                let characters = patch_varint_field(
                    &characters,
                    CHARACTER_FONT_NAME_NULL_FIELD,
                    current_font_null,
                    expected_font_null.then_some(1),
                )?;
                patch_length_delimited_field(
                    &characters,
                    CHARACTER_FONT_NAME_FIELD,
                    current_font_name,
                    expected_font_name.map(str::as_bytes),
                )
            },
        )?;
        let data = patch_varint_field(
            &data,
            PARAGRAPH_OVERRIDE_COUNT_FIELD,
            location.style.override_count.is_some(),
            Some(next_count),
        )?;
        object.replace_message(
            *message_index,
            RawMessage {
                type_: PARAGRAPH_STYLE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;
    let verified = locate_style(package, style_id)?;
    let actual = direct_overrides(&verified.style, &verified.message.data)?.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "legend paragraph style {style_id} stopped being an exact native variation"
        ))
    })?;
    if actual != *expected {
        return Err(Error::InvalidFormat(format!(
            "legend paragraph style {style_id} font wire patch failed validation"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn direct_index_patch_preserves_neighboring_and_unknown_fields() {
        const UNKNOWN_FIELD: u32 = 4_097;
        let mut generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultopacity: Some(0.8),
            ..Default::default()
        }
        .encode_to_vec();
        crate::wire::append_varint_field(&mut generated, UNKNOWN_FIELD, 42).unwrap();
        let mut original = tsch::LegendStyleArchive::default().encode_to_vec();
        crate::wire::append_length_delimited_field(
            &mut original,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();

        let direct = patch_direct_paragraph_style_index(&original, Some(3)).unwrap();
        assert_eq!(direct_paragraph_style_index(&direct).unwrap(), Some(3));
        let extension = generated_legend_style_extension(&direct).unwrap().unwrap();
        let decoded = tsch::generated::LegendStyleArchive::decode(extension).unwrap();
        assert_eq!(decoded.tschlegendmodeldefaultopacity, Some(0.8));
        assert!(
            crate::wire::parse_wire_fields(extension)
                .unwrap()
                .iter()
                .any(|field| field.number() == UNKNOWN_FIELD)
        );

        let inherited = patch_direct_paragraph_style_index(&direct, None).unwrap();
        assert_eq!(inherited, original);
    }

    #[test]
    fn negative_native_legend_paragraph_style_index_is_rejected() {
        let generated = tsch::generated::LegendStyleArchive {
            tschlegendmodeldefaultlabelparagraphstyleindex: Some(-1),
            ..Default::default()
        }
        .encode_to_vec();
        let mut data = tsch::LegendStyleArchive::default().encode_to_vec();
        crate::wire::append_length_delimited_field(
            &mut data,
            GENERATED_LEGEND_STYLE_EXTENSION_FIELD,
            &generated,
        )
        .unwrap();
        assert!(direct_paragraph_style_index(&data).is_err());
    }
}
