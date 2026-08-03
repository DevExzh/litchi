//! Typed, wire-preserving basic image adjustments for iWork image archives.

use prost::Message;

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

/// A normalized native image-adjustment amount.
///
/// iWork exposes exposure and saturation as percentages. Values are represented
/// here as a normalized multiplier in the inclusive range `-1.0..=1.0`, where
/// `0.25` corresponds to `25%` in the native inspector.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct ImageAdjustment(f32);

impl ImageAdjustment {
    /// The lowest native adjustment amount (`-100%`).
    pub const MINIMUM: Self = Self(-1.0);
    /// The neutral adjustment amount (`0%`).
    pub const NEUTRAL: Self = Self(0.0);
    /// The highest native adjustment amount (`100%`).
    pub const MAXIMUM: Self = Self(1.0);

    /// Construct one finite native image-adjustment amount.
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() || !(Self::MINIMUM.0..=Self::MAXIMUM.0).contains(&value) {
            return Err(Error::ParseError(
                "image adjustment must be finite and within -1.0..=1.0".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the normalized native amount.
    pub const fn as_f32(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for ImageAdjustment {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::new(value)
    }
}

/// State of iWork's native Enhance control.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ImageEnhancement {
    /// Leave automatic enhancement disabled.
    Disabled,
    /// Let iWork automatically enhance the image colors.
    Enabled,
}

impl ImageEnhancement {
    const fn from_native(value: bool) -> Self {
        if value { Self::Enabled } else { Self::Disabled }
    }

    const fn as_native(self) -> bool {
        matches!(self, Self::Enabled)
    }
}

/// The basic controls in iWork's Image inspector.
///
/// Optional fields preserve the native distinction between a control that was
/// never encoded and a control explicitly set to its neutral value. Advanced
/// native adjustment fields remain untouched by this focused API.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ImageAdjustments {
    /// Exposure adjustment, where `0.25` is `25%` in iWork.
    pub exposure: Option<ImageAdjustment>,
    /// Saturation adjustment, where `-1.0` produces grayscale output.
    pub saturation: Option<ImageAdjustment>,
    /// Optional state of iWork's automatic Enhance control.
    pub enhancement: Option<ImageEnhancement>,
}

impl ImageAdjustments {
    /// Construct adjustments with every basic control omitted.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            exposure: None,
            saturation: None,
            enhancement: None,
        }
    }

    /// Set or clear the optional exposure adjustment.
    #[must_use]
    pub const fn with_exposure(mut self, exposure: Option<ImageAdjustment>) -> Self {
        self.exposure = exposure;
        self
    }

    /// Set or clear the optional saturation adjustment.
    #[must_use]
    pub const fn with_saturation(mut self, saturation: Option<ImageAdjustment>) -> Self {
        self.saturation = saturation;
        self
    }

    /// Set or clear the optional automatic-enhancement state.
    #[must_use]
    pub const fn with_enhancement(mut self, enhancement: Option<ImageEnhancement>) -> Self {
        self.enhancement = enhancement;
        self
    }
}

pub(crate) fn image_adjustments_from_archive(
    image: &tsd::ImageArchive,
) -> Result<ImageAdjustments> {
    let Some(native) = image.image_adjustments.as_ref() else {
        return Ok(ImageAdjustments::default());
    };
    Ok(ImageAdjustments {
        exposure: adjustment_from_native(native.exposure, "exposure")?,
        saturation: adjustment_from_native(native.saturation, "saturation")?,
        enhancement: native.enhance.map(ImageEnhancement::from_native),
    })
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
        adjustments.exposure.map(|value| value.as_f32().to_bits()),
    )?;
    let data = patch_fixed32_field(
        &data,
        SATURATION_FIELD,
        native.saturation.is_some(),
        adjustments.saturation.map(|value| value.as_f32().to_bits()),
    )?;
    patch_varint_field(
        &data,
        ENHANCEMENT_FIELD,
        native.enhance.is_some(),
        adjustments
            .enhancement
            .map(ImageEnhancement::as_native)
            .map(u64::from),
    )
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
        let baseline = ImageAdjustments {
            exposure: Some(ImageAdjustment::NEUTRAL),
            saturation: Some(ImageAdjustment::NEUTRAL),
            enhancement: Some(ImageEnhancement::Disabled),
        };
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
            replacement.exposure
        );
        assert_eq!(
            adjustment_from_native(changed_native.saturation, "saturation").unwrap(),
            replacement.saturation
        );
        assert_eq!(
            changed_native.enhance.map(ImageEnhancement::from_native),
            replacement.enhancement
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
    fn no_op_adjustment_updates_reject_duplicate_raw_fields_transactionally() {
        let native = tsd::ImageAdjustmentsArchive {
            exposure: Some(0.25),
            saturation: Some(-0.5),
            enhance: Some(true),
            ..Default::default()
        };
        let requested = ImageAdjustments {
            exposure: Some(ImageAdjustment::new(0.25).unwrap()),
            saturation: Some(ImageAdjustment::new(-0.5).unwrap()),
            enhancement: Some(ImageEnhancement::Enabled),
        };

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
        duplicate_inner.extend(crate::varint::encode_varint(
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
        data.extend(crate::varint::encode_varint(u64::from(field_number) << 3));
        data.extend(crate::varint::encode_varint(value));
    }
}
