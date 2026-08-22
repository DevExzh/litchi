//! Typed, wire-preserving basic image adjustments for iWork image archives.

use litchi_iwa_common::shape::image::{
    Error as ImageAdjustmentError, ImageAdjustment, ImageAdjustments, ImageEnhancement,
};
use litchi_iwa_common::varint::decode_varint_from_bytes;
use litchi_iwa_common::{WireLimits, wire::WireDescent};

use crate::archive::RawMessage;
use crate::protobuf::tsd;
use crate::wire::{patch_fixed32_field, patch_length_delimited_field, patch_varint_field};
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

/// Decode only the image-adjustments edge from caller-owned ImageArchive
/// bytes.
///
/// The complete archive remains opaque.  This strict projection validates
/// canonical wire framing for every field it scans, rejects duplicate selected
/// fields, and retains the original nested payload so a subsequent rewrite can
/// preserve advanced/unknown fields byte-for-byte.
#[cfg(test)]
pub(crate) fn image_adjustments_from_bytes(source: &[u8]) -> Result<ImageAdjustments> {
    let mut budget = ImageWireBudget::new(source)?;
    let archive = inspect_image_archive(source, &mut budget, true)?;
    adjustments_from_raw(archive.adjustments)
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
        let mut budget = ImageWireBudget::new(original)?;
        let archive = inspect_image_archive(original, &mut budget, true).map_err(|error| {
            Error::InvalidFormat(format!(
                "{context} {image_id} image archive is invalid: {error}"
            ))
        })?;
        let current = adjustments_from_raw(archive.adjustments)?;
        if current == adjustments {
            return Ok(());
        }
        let has_adjustments = archive.adjustments.is_some();
        let current_payload = archive.adjustments.map_or(&[][..], |native| native.payload);
        let native = archive.adjustments.unwrap_or_default();
        let patched_adjustments =
            patch_raw_image_adjustments_payload(current_payload, native, adjustments, &mut budget)
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "{context} {image_id} image adjustment rewrite failed: {error}"
                    ))
                })?;
        let replacement =
            (!patched_adjustments.is_empty()).then_some(patched_adjustments.as_slice());
        let data = patch_length_delimited_field(
            original,
            IMAGE_ADJUSTMENTS_FIELD,
            has_adjustments,
            replacement,
        )?;
        budget.charge_rewrite_growth(original.len(), data.len())?;
        budget.ensure_output(data.len())?;
        let verified =
            inspect_image_archive(data.as_slice(), &mut budget, false).map_err(|error| {
                Error::InvalidFormat(format!(
                    "{context} {image_id} image adjustment readback failed: {error}"
                ))
            })?;
        if adjustments_from_raw(verified.adjustments)? != adjustments {
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

#[derive(Clone, Copy, Debug, Default)]
struct RawImageAdjustments<'source> {
    payload: &'source [u8],
    exposure: Option<f32>,
    saturation: Option<f32>,
    enhance: Option<bool>,
}

#[derive(Clone, Copy, Debug, Default)]
struct RawImageArchive<'source> {
    adjustments: Option<RawImageAdjustments<'source>>,
}

fn adjustments_from_raw(native: Option<RawImageAdjustments<'_>>) -> Result<ImageAdjustments> {
    let Some(native) = native else {
        return Ok(ImageAdjustments::default());
    };
    Ok(ImageAdjustments::new()
        .with_exposure(adjustment_from_native(native.exposure, "exposure")?)
        .with_saturation(adjustment_from_native(native.saturation, "saturation")?)
        .with_enhancement(native.enhance.map(image_enhancement_from_native)))
}

/// Rewrite the selected fields of a nested adjustment payload while leaving
/// every unselected source span (including advanced native controls) intact.
fn patch_raw_image_adjustments_payload(
    original: &[u8],
    native: RawImageAdjustments<'_>,
    adjustments: ImageAdjustments,
    budget: &mut ImageWireBudget,
) -> Result<Vec<u8>> {
    let data = patch_fixed32_field(
        original,
        EXPOSURE_FIELD,
        native.exposure.is_some(),
        adjustments.exposure().map(|value| value.value().to_bits()),
    )?;
    budget.ensure_output(data.len())?;
    let data = patch_fixed32_field(
        &data,
        SATURATION_FIELD,
        native.saturation.is_some(),
        adjustments
            .saturation()
            .map(|value| value.value().to_bits()),
    )?;
    budget.ensure_output(data.len())?;
    let data = patch_varint_field(
        &data,
        ENHANCEMENT_FIELD,
        native.enhance.is_some(),
        adjustments
            .enhancement()
            .map(image_enhancement_to_native)
            .map(u64::from),
    )?;
    budget.ensure_output(data.len())?;
    Ok(data)
}

/// Bounded source/work accounting for the focused raw-wire projection.
///
/// The common wire parser already performs fallible field storage and enforces
/// its hard input/field ceilings. This local budget charges the source archive
/// once for the logical rewrite and keeps aggregate input, field, output, and
/// growth accounting checked across the candidate and readback passes. The
/// candidate/readback scans intentionally do not charge the same source bytes
/// again: they validate one already-budgeted transaction rather than creating
/// additional rewrites.
#[derive(Debug)]
struct ImageWireBudget {
    limits: WireLimits,
    input_bytes: usize,
    fields: usize,
    work_bytes: usize,
}

impl ImageWireBudget {
    fn new(source: &[u8]) -> Result<Self> {
        let limits = WireLimits::default();
        let mut budget = Self {
            limits,
            input_bytes: 0,
            fields: 0,
            work_bytes: 0,
        };
        budget.charge_input(source.len())?;
        budget.charge_rewrite(source.len())?;
        Ok(budget)
    }

    fn charge_input(&mut self, amount: usize) -> Result<()> {
        self.input_bytes = self
            .input_bytes
            .checked_add(amount)
            .ok_or_else(|| Error::InvalidFormat("image archive input size overflow".to_owned()))?;
        if self.input_bytes > self.limits.max_input_bytes() {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::InputBytes,
                observed: self.input_bytes,
                limit: self.limits.max_input_bytes(),
            }));
        }
        Ok(())
    }

    fn charge_fields(&mut self, amount: usize) -> Result<()> {
        self.fields = self
            .fields
            .checked_add(amount)
            .ok_or_else(|| Error::InvalidFormat("image archive field count overflow".to_owned()))?;
        if self.fields > self.limits.max_fields() {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::Fields,
                observed: self.fields,
                limit: self.limits.max_fields(),
            }));
        }
        Ok(())
    }

    fn charge_rewrite(&mut self, amount: usize) -> Result<()> {
        self.work_bytes = self.work_bytes.checked_add(amount).ok_or_else(|| {
            Error::InvalidFormat("image archive rewrite work overflow".to_owned())
        })?;
        if self.work_bytes > self.limits.max_rewrite_work() {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::RewriteWork,
                observed: self.work_bytes,
                limit: self.limits.max_rewrite_work(),
            }));
        }
        Ok(())
    }

    fn charge_rewrite_growth(&mut self, source: usize, output: usize) -> Result<()> {
        let growth = output.saturating_sub(source);
        self.charge_rewrite(growth)
    }

    fn ensure_output(&self, amount: usize) -> Result<()> {
        if amount > self.limits.max_output_bytes() {
            return Err(Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: litchi_iwa_common::LimitKind::OutputBytes,
                observed: amount,
                limit: self.limits.max_output_bytes(),
            }));
        }
        Ok(())
    }
}

fn inspect_image_archive<'source>(
    source: &'source [u8],
    budget: &mut ImageWireBudget,
    root_already_charged: bool,
) -> Result<RawImageArchive<'source>> {
    budget.ensure_output(source.len())?;
    let limits = budget.limits;
    let preflight =
        litchi_iwa_common::wire::preflight_wire_tree_with_limits(source, limits, |visit| {
            let field = visit.field();
            field.validate_canonical_framing()?;
            if field.wire_type() == 0 {
                let payload = field.payload();
                let (_, encoded) = decode_varint_from_bytes(payload).map_err(|error| {
                    litchi_iwa_common::Error::InvalidFormat(format!(
                        "protobuf varint value is invalid: {error}"
                    ))
                })?;
                if encoded != payload.len() {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "protobuf varint value is not canonical".to_owned(),
                    ));
                }
            }
            if visit.path().is_empty() && field.number() == IMAGE_ADJUSTMENTS_FIELD {
                Ok(WireDescent::Descend)
            } else {
                Ok(WireDescent::Skip)
            }
        })?;
    let scanned_bytes = preflight.scanned_bytes();
    let additional_input = if root_already_charged {
        scanned_bytes.checked_sub(source.len()).ok_or_else(|| {
            Error::InvalidFormat("image archive preflight byte accounting underflow".to_owned())
        })?
    } else {
        scanned_bytes
    };
    budget.charge_input(additional_input)?;
    budget.charge_fields(preflight.fields())?;

    let view = litchi_iwa_common::wire::parse_wire_view_with_limits(source, limits)?;
    let mut adjustments = None;
    for field in view.fields() {
        if field.number() != IMAGE_ADJUSTMENTS_FIELD {
            continue;
        }
        if adjustments.is_some() {
            return Err(Error::InvalidFormat(
                "image archive has duplicate image-adjustments fields".to_owned(),
            ));
        }
        if field.wire_type() != 2 {
            return Err(Error::InvalidFormat(
                "image-adjustments field is not length-delimited".to_owned(),
            ));
        }
        let payload = field.payload();
        adjustments = Some(inspect_image_adjustments(payload)?);
    }
    Ok(RawImageArchive { adjustments })
}

fn inspect_image_adjustments<'source>(
    source: &'source [u8],
) -> Result<RawImageAdjustments<'source>> {
    let limits = WireLimits::default();
    let view = litchi_iwa_common::wire::parse_wire_view_with_limits(source, limits)?;
    let mut exposure = None;
    let mut saturation = None;
    let mut enhance = None;
    for field in view.fields() {
        match field.number() {
            EXPOSURE_FIELD => {
                if exposure.is_some() {
                    return Err(Error::InvalidFormat(
                        "image adjustments has duplicate exposure fields".to_owned(),
                    ));
                }
                if field.wire_type() != 5 {
                    return Err(Error::InvalidFormat(
                        "image adjustments exposure is not fixed32".to_owned(),
                    ));
                }
                let payload = field.payload();
                let bytes: [u8; 4] = payload.try_into().map_err(|_error| {
                    Error::InvalidFormat("image adjustments exposure is truncated".to_owned())
                })?;
                exposure = Some(f32::from_bits(u32::from_le_bytes(bytes)));
            },
            SATURATION_FIELD => {
                if saturation.is_some() {
                    return Err(Error::InvalidFormat(
                        "image adjustments has duplicate saturation fields".to_owned(),
                    ));
                }
                if field.wire_type() != 5 {
                    return Err(Error::InvalidFormat(
                        "image adjustments saturation is not fixed32".to_owned(),
                    ));
                }
                let payload = field.payload();
                let bytes: [u8; 4] = payload.try_into().map_err(|_error| {
                    Error::InvalidFormat("image adjustments saturation is truncated".to_owned())
                })?;
                saturation = Some(f32::from_bits(u32::from_le_bytes(bytes)));
            },
            ENHANCEMENT_FIELD => {
                if enhance.is_some() {
                    return Err(Error::InvalidFormat(
                        "image adjustments has duplicate enhancement fields".to_owned(),
                    ));
                }
                if field.wire_type() != 0 {
                    return Err(Error::InvalidFormat(
                        "image adjustments enhancement is not varint".to_owned(),
                    ));
                }
                let (value, encoded) =
                    decode_varint_from_bytes(field.payload()).map_err(|error| {
                        Error::InvalidFormat(format!(
                            "image adjustments enhancement is invalid: {error}"
                        ))
                    })?;
                if encoded != field.payload().len() {
                    return Err(Error::InvalidFormat(
                        "image adjustments enhancement is not canonical".to_owned(),
                    ));
                }
                enhance = Some(match value {
                    0 => false,
                    1 => true,
                    _ => {
                        return Err(Error::InvalidFormat(
                            "image adjustments enhancement is not a canonical bool".to_owned(),
                        ));
                    },
                });
            },
            _ => {},
        }
    }
    Ok(RawImageAdjustments {
        payload: source,
        exposure,
        saturation,
        enhance,
    })
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
    use prost::Message as _;

    const TEST_ARCHIVE_NAME: &str = "Index/Image.iwa";

    fn patch_image_adjustments_payload(
        original: &[u8],
        native: &tsd::ImageAdjustmentsArchive,
        adjustments: ImageAdjustments,
    ) -> Result<Vec<u8>> {
        let mut budget = ImageWireBudget::new(original)?;
        patch_raw_image_adjustments_payload(
            original,
            RawImageAdjustments {
                payload: original,
                exposure: native.exposure,
                saturation: native.saturation,
                enhance: native.enhance,
            },
            adjustments,
            &mut budget,
        )
    }

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
    fn raw_projection_preserves_unknown_outer_and_inner_fields() {
        let native = tsd::ImageAdjustmentsArchive {
            exposure: Some(0.0),
            saturation: Some(0.0),
            contrast: Some(0.4),
            enhance: Some(false),
            ..Default::default()
        };
        let mut payload = native.encode_to_vec();
        append_unknown_varint(&mut payload, 99, 990);
        let mut source = tsd::ImageArchive::default().encode_to_vec();
        crate::wire::append_length_delimited_field(&mut source, IMAGE_ADJUSTMENTS_FIELD, &payload)
            .unwrap();
        append_unknown_varint(&mut source, 100, 1_000);
        let before = source.clone();

        let decoded = image_adjustments_from_bytes(&source).unwrap();
        assert_eq!(decoded.exposure(), Some(ImageAdjustment::NEUTRAL));
        assert_eq!(decoded.saturation(), Some(ImageAdjustment::NEUTRAL));
        assert_eq!(decoded.enhancement(), Some(ImageEnhancement::Disabled));

        let mut package = package_with_image(source);
        let replacement = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        replace_image_adjustments(
            &mut package,
            TEST_ARCHIVE_NAME,
            1,
            "test image",
            replacement,
        )
        .unwrap();
        let changed = image_message(&package);
        assert!(
            changed
                .windows(4)
                .any(|window| window == [0x98, 0x06, 0xde, 0x07])
        );
        assert!(
            changed
                .windows(4)
                .any(|window| window == [0xa0, 0x06, 0xe8, 0x07])
        );
        assert_ne!(changed, before.as_slice());

        let reset = ImageAdjustments::default();
        replace_image_adjustments(&mut package, TEST_ARCHIVE_NAME, 1, "test image", reset).unwrap();
        let reset_bytes = image_message(&package);
        assert!(
            reset_bytes
                .windows(4)
                .any(|window| window == [0x98, 0x06, 0xde, 0x07])
        );
        assert!(
            reset_bytes
                .windows(4)
                .any(|window| window == [0xa0, 0x06, 0xe8, 0x07])
        );
        assert_eq!(image_adjustments_from_bytes(&reset_bytes).unwrap(), reset);
    }

    #[test]
    fn raw_projection_rejects_malformed_selected_framing() {
        let malformed = [
            vec![0x72, 0x02, 0x09, 0x00], // field 14, exposure has wrong wire type
            vec![0x72, 0x02, 0x8, 0x80],  // selected varint is truncated
            vec![0xf0, 0x80, 0x00, 0x00], // overlong unknown field key
        ];
        for source in malformed {
            assert!(
                image_adjustments_from_bytes(&source).is_err(),
                "malformed source unexpectedly decoded: {source:?}"
            );
        }
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

        let removed =
            patch_length_delimited_field(encoded.as_slice(), IMAGE_ADJUSTMENTS_FIELD, true, None)
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

    #[test]
    fn opaque_archive_at_rewrite_work_boundary_is_accepted() {
        let native = tsd::ImageAdjustmentsArchive {
            exposure: Some(0.0),
            saturation: Some(0.0),
            enhance: Some(false),
            ..Default::default()
        };
        let mut source = tsd::ImageArchive::default().encode_to_vec();
        crate::wire::append_length_delimited_field(
            &mut source,
            IMAGE_ADJUSTMENTS_FIELD,
            native.encode_to_vec().as_slice(),
        )
        .unwrap();

        let unknown_key = (u64::from(100u32) << 3) | 2;
        let key_bytes = litchi_iwa_common::varint::encoded_len(unknown_key);
        let mut unknown_len = WireLimits::MAX_REWRITE_WORK
            .checked_sub(source.len())
            .and_then(|remaining| remaining.checked_sub(key_bytes + 4))
            .expect("boundary fixture should leave room for an opaque field");
        for _ in 0..4 {
            unknown_len = WireLimits::MAX_REWRITE_WORK
                .checked_sub(source.len())
                .and_then(|remaining| {
                    remaining.checked_sub(
                        key_bytes + litchi_iwa_common::varint::encoded_len(unknown_len as u64),
                    )
                })
                .expect("opaque field length should remain in range");
        }
        crate::wire::append_length_delimited_field(
            &mut source,
            100,
            vec![0; unknown_len].as_slice(),
        )
        .unwrap();
        assert_eq!(source.len(), WireLimits::MAX_REWRITE_WORK);

        let requested = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        let mut package = package_with_image(source);
        replace_image_adjustments(&mut package, TEST_ARCHIVE_NAME, 1, "test image", requested)
            .unwrap();
        let changed = image_message(&package);
        assert_eq!(changed.len(), WireLimits::MAX_REWRITE_WORK);
        assert_eq!(image_adjustments_from_bytes(&changed).unwrap(), requested);
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

    fn image_message(package: &IWorkPackage) -> Vec<u8> {
        package
            .archive(TEST_ARCHIVE_NAME)
            .unwrap()
            .object(1)
            .unwrap()
            .messages
            .iter()
            .find(|message| message.type_ == IMAGE_ARCHIVE_MESSAGE_TYPE)
            .unwrap()
            .data
            .clone()
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(field_number) << 3,
        ));
        data.extend(litchi_iwa_common::varint::encode_varint(value));
    }
}
