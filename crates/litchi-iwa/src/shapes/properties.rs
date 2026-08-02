//! Typed, wire-preserving properties for iWork drawables.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::{tsd, tswp};
use crate::wire::{
    patch_length_delimited_field, patch_varint_field, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_INFO_ARCHIVE_SUPER_FIELD: u32 = 1;
const SHAPE_ARCHIVE_SUPER_FIELD: u32 = 1;
const WRAPPED_DRAWABLE_ARCHIVE_SUPER_FIELD: u32 = 1;
const DRAWABLE_HYPERLINK_URL_FIELD: u32 = 4;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const DRAWABLE_ASPECT_RATIO_LOCKED_FIELD: u32 = 7;
const DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD: u32 = 8;

/// Properties stored directly on an iWork drawable.
///
/// Optional fields retain the distinction between an absent protobuf field and
/// an explicit default. Native editors can disable some properties for
/// particular drawable types even though the fields remain part of the shared
/// archive.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DrawableProperties {
    pub hyperlink_url: Option<String>,
    pub locked: Option<bool>,
    pub aspect_ratio_locked: Option<bool>,
    pub accessibility_description: Option<String>,
}

pub(crate) fn shape_properties(
    package: &IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
) -> Result<DrawableProperties> {
    let archive = package.archive(archive_name)?;
    let object = archive.object(drawable_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork drawable object {drawable_id} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    if messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork drawable {drawable_id} must have exactly one ShapeInfo payload"
        )));
    }
    let shape = tswp::ShapeInfoArchive::decode(messages[0].data.as_slice())?;
    Ok(properties_from_shape(&shape))
}

pub(crate) fn set_shape_properties(
    package: &mut IWorkPackage,
    archive_name: &str,
    drawable_id: u64,
    replacement: &DrawableProperties,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(drawable_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork drawable object {drawable_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        if indexes.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork drawable {drawable_id} must have exactly one ShapeInfo payload"
            )));
        }

        let message_index = indexes[0];
        let original = object.messages[message_index].data.as_slice();
        let shape = tswp::ShapeInfoArchive::decode(original)?;
        let current = properties_from_shape(&shape);
        let data = transform_length_delimited_field(
            original,
            SHAPE_INFO_ARCHIVE_SUPER_FIELD,
            |shape_archive| {
                transform_length_delimited_field(
                    shape_archive,
                    SHAPE_ARCHIVE_SUPER_FIELD,
                    |drawable_archive| {
                        patch_drawable_properties(drawable_archive, &current, replacement)
                    },
                )
            },
        )?;
        let verified = tswp::ShapeInfoArchive::decode(data.as_slice())?;
        if properties_from_shape(&verified) != *replacement {
            return Err(Error::InvalidFormat(
                "iWork drawable properties patch failed validation".to_owned(),
            ));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: SHAPE_INFO_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn properties_from_shape(shape: &tswp::ShapeInfoArchive) -> DrawableProperties {
    drawable_properties(&shape.super_.super_)
}

/// Extract shared user-facing properties from a decoded drawable archive.
pub(crate) fn drawable_properties(drawable: &tsd::DrawableArchive) -> DrawableProperties {
    DrawableProperties {
        hyperlink_url: drawable.hyperlink_url.clone(),
        locked: drawable.locked,
        aspect_ratio_locked: drawable.aspect_ratio_locked,
        accessibility_description: drawable.accessibility_description.clone(),
    }
}

/// Patch only the shared property fields of one raw `TSD.DrawableArchive`.
///
/// The operation retains all unknown fields and preserves the wire distinction
/// between absent optional fields and explicitly encoded defaults.
pub(crate) fn patch_drawable_properties(
    data: &[u8],
    current: &DrawableProperties,
    replacement: &DrawableProperties,
) -> Result<Vec<u8>> {
    let mut data = patch_length_delimited_field(
        data,
        DRAWABLE_HYPERLINK_URL_FIELD,
        current.hyperlink_url.is_some(),
        replacement.hyperlink_url.as_deref().map(str::as_bytes),
    )?;
    data = patch_varint_field(
        &data,
        DRAWABLE_LOCKED_FIELD,
        current.locked.is_some(),
        replacement.locked.map(u64::from),
    )?;
    data = patch_varint_field(
        &data,
        DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
        current.aspect_ratio_locked.is_some(),
        replacement.aspect_ratio_locked.map(u64::from),
    )?;
    patch_length_delimited_field(
        &data,
        DRAWABLE_ACCESSIBILITY_DESCRIPTION_FIELD,
        current.accessibility_description.is_some(),
        replacement
            .accessibility_description
            .as_deref()
            .map(str::as_bytes),
    )
}

/// Patch a message whose first field embeds a `TSD.DrawableArchive`.
///
/// `TSD.ImageArchive` and `TSD.MovieArchive` use this shape. Keeping the
/// wrapper handling here lets each editor validate ownership while sharing the
/// lossless field-patching implementation.
pub(crate) fn patch_wrapped_drawable_properties(
    data: &[u8],
    current: &DrawableProperties,
    replacement: &DrawableProperties,
) -> Result<Vec<u8>> {
    transform_length_delimited_field(data, WRAPPED_DRAWABLE_ARCHIVE_SUPER_FIELD, |drawable| {
        patch_drawable_properties(drawable, current, replacement)
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protobuf::tsd;

    #[test]
    fn properties_patch_preserves_unknowns_and_restores_exactly() {
        let drawable = tsd::DrawableArchive {
            hyperlink_url: Some("https://example.test/original".to_owned()),
            locked: Some(false),
            aspect_ratio_locked: Some(false),
            accessibility_description: Some("Original description".to_owned()),
            ..Default::default()
        };
        let mut original = drawable.encode_to_vec();
        append_unknown_varint(&mut original, 99, 990);
        let baseline = DrawableProperties {
            hyperlink_url: drawable.hyperlink_url,
            locked: drawable.locked,
            aspect_ratio_locked: drawable.aspect_ratio_locked,
            accessibility_description: drawable.accessibility_description,
        };
        let replacement = DrawableProperties {
            hyperlink_url: Some("keynote://litchi/example".to_owned()),
            locked: Some(true),
            aspect_ratio_locked: Some(true),
            accessibility_description: Some("Accessible 文本 box".to_owned()),
        };

        let changed = patch_drawable_properties(&original, &baseline, &replacement).unwrap();
        assert_eq!(
            tsd::DrawableArchive::decode(changed.as_slice())
                .unwrap()
                .hyperlink_url,
            replacement.hyperlink_url
        );
        assert!(
            changed
                .windows(3)
                .any(|window| window == [0x98, 0x06, 0xde])
        );

        let restored = patch_drawable_properties(&changed, &replacement, &baseline).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn wrapped_properties_patch_preserves_outer_and_drawable_unknowns() {
        let image = tsd::ImageArchive {
            super_: tsd::DrawableArchive {
                locked: Some(false),
                accessibility_description: Some("Original image".to_owned()),
                ..Default::default()
            },
            ..Default::default()
        };
        let baseline = drawable_properties(&image.super_);
        let replacement = DrawableProperties {
            locked: Some(true),
            accessibility_description: Some("Accessible image".to_owned()),
            ..Default::default()
        };
        let mut original = image.encode_to_vec();
        append_unknown_varint(&mut original, 99, 123);

        let changed =
            patch_wrapped_drawable_properties(&original, &baseline, &replacement).unwrap();
        let decoded = tsd::ImageArchive::decode(changed.as_slice()).unwrap();
        assert_eq!(drawable_properties(&decoded.super_), replacement);
        assert!(
            changed
                .windows(3)
                .any(|window| window == [0x98, 0x06, 0x7b])
        );

        let cleared = patch_wrapped_drawable_properties(
            &changed,
            &replacement,
            &DrawableProperties::default(),
        )
        .unwrap();
        let cleared_image = tsd::ImageArchive::decode(cleared.as_slice()).unwrap();
        assert_eq!(
            drawable_properties(&cleared_image.super_),
            DrawableProperties::default()
        );

        let restored =
            patch_wrapped_drawable_properties(&cleared, &DrawableProperties::default(), &baseline)
                .unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn properties_patch_rejects_duplicate_and_malformed_fields() {
        let drawable = tsd::DrawableArchive {
            locked: Some(false),
            ..Default::default()
        };
        let current = DrawableProperties {
            locked: Some(false),
            ..Default::default()
        };
        let replacement = DrawableProperties {
            locked: Some(true),
            ..Default::default()
        };

        let mut duplicate = drawable.encode_to_vec();
        duplicate.extend(crate::varint::encode_varint(5 << 3));
        duplicate.push(1);
        assert!(patch_drawable_properties(&duplicate, &current, &replacement).is_err());

        let mut malformed = drawable.encode_to_vec();
        malformed.extend(crate::varint::encode_varint((8 << 3) | 2));
        malformed.push(3);
        malformed.push(b'x');
        assert!(patch_drawable_properties(&malformed, &current, &replacement).is_err());
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
        data.extend(crate::varint::encode_varint(value));
    }
}
