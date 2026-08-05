//! Typed, wire-preserving basic image adjustments for iWork image archives.

use prost::Message;

use litchi_iwa_common::shape::image::{
    Error as ImageAdjustmentError, ImageAdjustment, ImageAdjustments, ImageEnhancement,
};

use crate::archive::RawMessage;
use crate::protobuf::tsd;
use crate::wire::{
    patch_fixed32_field, patch_length_delimited_field, patch_varint_field,
    repeated_length_delimited_payloads,
};
use crate::{Error, IWorkPackage, Result};

const IMAGE_ARCHIVE_MESSAGE_TYPE: u32 = 3_005;
const IMAGE_ADJUSTMENTS_FIELD: u32 = 14;
const EXPOSURE_FIELD: u32 = 1;
const SATURATION_FIELD: u32 = 2;
const ENHANCEMENT_FIELD: u32 = 13;

impl From<ImageAdjustmentError> for crate::Error {
    fn from(error: ImageAdjustmentError) -> Self {
        Self::ParseError(error.to_string())
    }
}

pub(crate) fn image_adjustments_from_archive(
    image: &tsd::ImageArchive,
) -> Result<ImageAdjustments> {
    let Some(native) = image.image_adjustments.as_ref() else {
        return Ok(ImageAdjustments::default());
    };
    Ok(ImageAdjustments::new()
        .with_exposure(adjustment_from_native(native.exposure, "exposure")?)
        .with_saturation(adjustment_from_native(native.saturation, "saturation")?)
        .with_enhancement(native.enhance.map(image_enhancement_from_native)))
}

pub(crate) fn replace_image_adjustments(
    package: &mut IWorkPackage,
    archive_name: &str,
    image_id: u64,
    context: &str,
    adjustments: ImageAdjustments,
) -> Result<ImageAdjustments> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(image_id).ok_or_else(|| {
            Error::InvalidFormat(format!("{context} object {image_id} is missing"))
        })?;
        let message_indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == IMAGE_ARCHIVE_MESSAGE_TYPE).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = message_indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "{context} {image_id} must have exactly one ImageArchive payload"
            )));
        };
        let message_index = *message_index;
        let original = object.messages[message_index].data.as_slice();
        let image = tsd::ImageArchive::decode(original)?;
        let current = image_adjustments_from_archive(&image)?;
        let has_adjustments = image.image_adjustments.is_some();
        let native = image.image_adjustments.unwrap_or_default();
        let payloads = repeated_length_delimited_payloads(original, IMAGE_ADJUSTMENTS_FIELD)?;
        let current_payload = match (has_adjustments, payloads.as_slice()) {
            (false, []) => &[][..],
            (true, [payload]) => *payload,
            (true, []) => {
                return Err(Error::InvalidFormat(format!(
                    "{context} {image_id} has image adjustments without a raw payload"
                )));
            },
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "{context} {image_id} must have exactly one image-adjustments payload"
                )));
            },
        };
        let patched_adjustments =
            patch_image_adjustments_payload(current_payload, &native, adjustments)?;
        if current == adjustments {
            return Ok(());
        }
        let replacement =
            (!patched_adjustments.is_empty()).then_some(patched_adjustments.as_slice());
        let data = patch_length_delimited_field(
            original,
            IMAGE_ADJUSTMENTS_FIELD,
            image.image_adjustments.is_some(),
            replacement,
        )?;
        let verified = tsd::ImageArchive::decode(data.as_slice())?;
        if image_adjustments_from_archive(&verified)? != adjustments {
            return Err(Error::InvalidFormat(format!(
                "{context} image adjustment patch failed validation"
            )));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: IMAGE_ARCHIVE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;
    Ok(adjustments)
}

fn adjustment_from_native(value: Option<f32>, control: &str) -> Result<Option<ImageAdjustment>> {
    value
        .map(ImageAdjustment::new)
        .transpose()
        .map_err(|error| {
            Error::InvalidFormat(format!("image {control} adjustment is invalid: {error}"))
        })
}

fn patch_image_adjustments_payload(
    original: &[u8],
    native: &tsd::ImageAdjustmentsArchive,
    adjustments: ImageAdjustments,
) -> Result<Vec<u8>> {
    let data = patch_fixed32_field(
        original,
        EXPOSURE_FIELD,
        native.exposure.is_some(),
        adjustments.exposure().map(|value| value.value().to_bits()),
    )?;
    let data = patch_fixed32_field(
        &data,
        SATURATION_FIELD,
        native.saturation.is_some(),
        adjustments.saturation().map(|value| value.value().to_bits()),
    )?;
    patch_varint_field(
        &data,
        ENHANCEMENT_FIELD,
        native.enhance.is_some(),
        adjustments
            .enhancement()
            .map(image_enhancement_to_native)
            .map(u64::from),
    )
}

const fn image_enhancement_from_native(value: bool) -> ImageEnhancement {
    if value {
        ImageEnhancement::Enabled
    } else {
        ImageEnhancement::Disabled
    }
}

const fn image_enhancement_to_native(value: ImageEnhancement) -> bool {
    matches!(value, ImageEnhancement::Enabled)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject};

    const TEST_ARCHIVE_NAME: &str = "Index/Image.iwa";

    #[test]
    fn adjustment_values_reject_invalid_native_percentages() {
        for invalid in [f32::NAN, f32::INFINITY, -1.01, 1.01] {
            assert!(ImageAdjustment::new(invalid).is_err());
        }
        assert_eq!(
            ImageAdjustment::new(-1.0).unwrap(),
            ImageAdjustment::MINIMUM
        );
        assert_eq!(ImageAdjustment::new(1.0).unwrap(), ImageAdjustment::MAXIMUM);
    }

    #[test]
    fn basic_adjustment_patch_preserves_unknown_advanced_fields() {
        let native = tsd::ImageAdjustmentsArchive {
            exposure: Some(0.0),
            saturation: Some(0.0),
            contrast: Some(0.4),
            enhance: Some(false),
            ..Default::default()
        };
        let baseline = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::NEUTRAL))
            .with_saturation(Some(ImageAdjustment::NEUTRAL))
            .with_enhancement(Some(ImageEnhancement::Disabled));
        let replacement = ImageAdjustments::default()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        let mut original = native.encode_to_vec();
        append_unknown_varint(&mut original, 99, 990);

        let changed = patch_image_adjustments_payload(&original, &native, replacement).unwrap();
        let changed_native = tsd::ImageAdjustmentsArchive::decode(changed.as_slice()).unwrap();
        assert_eq!(
            adjustment_from_native(changed_native.exposure, "exposure").unwrap(),
            replacement.exposure()
        );
        assert_eq!(
            adjustment_from_native(changed_native.saturation, "saturation").unwrap(),
            replacement.saturation()
        );
        assert_eq!(
            changed_native.enhance.map(image_enhancement_from_native),
            replacement.enhancement()
        );
        assert_eq!(changed_native.contrast, native.contrast);
        assert!(
            changed
                .windows(3)
                .any(|window| window == [0x98, 0x06, 0xde])
        );

        let restored =
            patch_image_adjustments_payload(&changed, &changed_native, baseline).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn archive_adapter_preserves_omitted_and_explicit_neutral_presence() {
        let omitted = tsd::ImageArchive::default();
        let omitted_adjustments = image_adjustments_from_archive(&omitted).unwrap();
        assert_eq!(omitted_adjustments.exposure(), None);
        assert_eq!(omitted_adjustments.saturation(), None);
        assert_eq!(omitted_adjustments.enhancement(), None);

        let explicit = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::NEUTRAL))
            .with_saturation(Some(ImageAdjustment::NEUTRAL))
            .with_enhancement(Some(ImageEnhancement::Disabled));
        let payload = patch_image_adjustments_payload(
            &[],
            &tsd::ImageAdjustmentsArchive::default(),
            explicit,
        )
        .unwrap();
        let encoded = patch_length_delimited_field(
            &[],
            IMAGE_ADJUSTMENTS_FIELD,
            false,
            Some(payload.as_slice()),
        )
        .unwrap();
        let decoded = tsd::ImageArchive::decode(encoded.as_slice()).unwrap();
        let decoded_adjustments = image_adjustments_from_archive(&decoded).unwrap();
        assert_eq!(decoded_adjustments, explicit);
        assert_ne!(decoded_adjustments, omitted_adjustments);

        let removed = patch_length_delimited_field(
            encoded.as_slice(),
            IMAGE_ADJUSTMENTS_FIELD,
            true,
            None,
        )
        .unwrap();
        assert_eq!(
            image_adjustments_from_archive(&tsd::ImageArchive::decode(removed.as_slice()).unwrap())
                .unwrap(),
            omitted_adjustments
        );
    }

    #[test]
    fn no_op_adjustment_updates_reject_duplicate_raw_fields_transactionally() {
        let native = tsd::ImageAdjustmentsArchive {
            exposure: Some(0.25),
            saturation: Some(-0.5),
            enhance: Some(true),
            ..Default::default()
        };
        let requested = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));

        let mut duplicate_outer = tsd::ImageArchive::default().encode_to_vec();
        let payload = native.encode_to_vec();
        crate::wire::append_length_delimited_field(
            &mut duplicate_outer,
            IMAGE_ADJUSTMENTS_FIELD,
            &payload,
        )
        .unwrap();
        crate::wire::append_length_delimited_field(
            &mut duplicate_outer,
            IMAGE_ADJUSTMENTS_FIELD,
            &payload,
        )
        .unwrap();
        assert_duplicate_rejected(duplicate_outer, requested);

        let mut duplicate_inner = native.encode_to_vec();
        duplicate_inner.extend(litchi_iwa_common::varint::encode_varint(
            (u64::from(EXPOSURE_FIELD) << 3) | 5,
        ));
        duplicate_inner.extend(native.exposure.unwrap().to_bits().to_le_bytes());
        let mut image = tsd::ImageArchive::default().encode_to_vec();
        crate::wire::append_length_delimited_field(
            &mut image,
            IMAGE_ADJUSTMENTS_FIELD,
            &duplicate_inner,
        )
        .unwrap();
        assert_duplicate_rejected(image, requested);
    }

    fn assert_duplicate_rejected(data: Vec<u8>, requested: ImageAdjustments) {
        let mut package = package_with_image(data);
        let before = package.entry(TEST_ARCHIVE_NAME).unwrap().to_vec();
        assert!(
            replace_image_adjustments(&mut package, TEST_ARCHIVE_NAME, 1, "test image", requested,)
                .is_err()
        );
        assert_eq!(package.entry(TEST_ARCHIVE_NAME).unwrap(), before.as_slice());
    }

    fn package_with_image(data: Vec<u8>) -> IWorkPackage {
        let archive = Archive {
            objects: vec![
                ArchiveObject::new(
                    1,
                    vec![RawMessage {
                        type_: IMAGE_ARCHIVE_MESSAGE_TYPE,
                        data,
                    }],
                )
                .unwrap(),
            ],
        };
        let mut package = IWorkPackage::new();
        package
            .replace_archive(TEST_ARCHIVE_NAME, &archive)
            .unwrap();
        package
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(field_number) << 3,
        ));
        data.extend(litchi_iwa_common::varint::encode_varint(value));
    }
}
