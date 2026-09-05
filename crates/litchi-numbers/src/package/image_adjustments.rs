//! Bounded Numbers image-adjustment projection and source-preserving rewrite.
//!
//! This module is the private native bridge for the legacy `litchi-iwa` host.
//! The complete `TSD.ImageArchive` payload remains byte-authoritative in the
//! host; this focused owner projects only the three semantic Image inspector
//! controls and delegates wire validation and rewriting to the strict Buffa
//! codec.

use litchi_iwa_common::{
    WireLimits,
    shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement},
};
use litchi_iwa_protos::image_adjustments_codec::{
    self as codec, ImageAdjustmentsSnapshot, ImageAdjustmentsWrite,
};
use thiserror::Error;

/// Typed failures at the hidden Numbers image-adjustment seam.
#[doc(hidden)]
#[derive(Debug, Clone, PartialEq, Eq, Error)]
pub enum ImageAdjustmentsError {
    /// The bounded Buffa/wire projection rejected the source or candidate.
    #[error("image-adjustments codec rejected the payload: {0}")]
    Codec(#[from] codec::DecodeError),
    /// The projected values failed the common semantic boundary.
    #[error("invalid image adjustment value: {0}")]
    Semantic(#[from] litchi_iwa_common::shape::image::Error),
}

fn codec_options(source: &[u8], limits: WireLimits) -> codec::DecodeOptions {
    let recursion_limit = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    codec::DecodeOptions::new(
        limits.max_input_bytes().min(source.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion_limit,
    )
    .with_max_output_bytes(limits.max_output_bytes())
}

/// Decode one borrowed Numbers ImageArchive payload into semantic controls.
#[doc(hidden)]
pub fn __decode_image_adjustments_payload(
    source: &[u8],
    limits: WireLimits,
) -> Result<ImageAdjustments, ImageAdjustmentsError> {
    let snapshot = codec::decode_image_adjustments(source, codec_options(source, limits))?;
    adjustments_from_snapshot(snapshot)
}

/// Rewrite one borrowed Numbers ImageArchive payload while preserving every
/// unknown field and the source's optional-field presence shape.
#[doc(hidden)]
pub fn __rewrite_image_adjustments_payload(
    source: &[u8],
    adjustments: ImageAdjustments,
    limits: WireLimits,
) -> Result<Vec<u8>, ImageAdjustmentsError> {
    let write = write_from_adjustments(adjustments);
    Ok(codec::rewrite_image_adjustments(
        source,
        write,
        codec_options(source, limits),
    )?)
}

fn adjustments_from_snapshot(
    snapshot: ImageAdjustmentsSnapshot<'_>,
) -> Result<ImageAdjustments, ImageAdjustmentsError> {
    Ok(ImageAdjustments::new()
        .with_exposure(snapshot.exposure().map(ImageAdjustment::new).transpose()?)
        .with_saturation(
            snapshot
                .saturation()
                .map(ImageAdjustment::new)
                .transpose()?,
        )
        .with_enhancement(snapshot.enhance().map(image_enhancement_from_native)))
}

fn write_from_adjustments(adjustments: ImageAdjustments) -> ImageAdjustmentsWrite {
    ImageAdjustmentsWrite::from_values(
        adjustments.exposure().map(ImageAdjustment::value),
        adjustments.saturation().map(ImageAdjustment::value),
        adjustments.enhancement().map(image_enhancement_to_native),
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
    use litchi_iwa_common::WireLimits;
    use litchi_iwa_common::shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement};

    use super::{__decode_image_adjustments_payload, __rewrite_image_adjustments_payload};

    #[test]
    fn focused_bridge_maps_native_controls_and_preserves_unknown_bytes() {
        let source = [
            0x0a, 0x00, // ImageArchive.super
            0x72, 0x10, // ImageArchive.imageAdjustments, length 16
            0x0d, 0x00, 0x00, 0x00, 0x00, // exposure = 0.0
            0x15, 0x00, 0x00, 0x00, 0x00, // saturation = 0.0
            0x68, 0x00, // enhance = false
            0x98, 0x06, 0xde, 0x07, // unknown nested field
            0xa0, 0x06, 0xe8, 0x07, // unknown outer field
        ];
        let baseline = __decode_image_adjustments_payload(&source, WireLimits::default()).unwrap();
        assert_eq!(baseline.exposure(), Some(ImageAdjustment::NEUTRAL));
        assert_eq!(baseline.saturation(), Some(ImageAdjustment::NEUTRAL));
        assert_eq!(baseline.enhancement(), Some(ImageEnhancement::Disabled));

        let replacement = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        let changed =
            __rewrite_image_adjustments_payload(&source, replacement, WireLimits::default())
                .unwrap();
        assert_eq!(
            __decode_image_adjustments_payload(&changed, WireLimits::default()).unwrap(),
            replacement
        );
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

        let restored =
            __rewrite_image_adjustments_payload(&changed, baseline, WireLimits::default()).unwrap();
        assert_eq!(restored, source);
    }

    #[test]
    fn focused_bridge_preserves_omitted_outer_and_inner_controls() {
        let source = [0x0a, 0x00];
        let baseline = __decode_image_adjustments_payload(&source, WireLimits::default()).unwrap();
        assert_eq!(baseline, ImageAdjustments::default());
        let explicit = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::NEUTRAL))
            .with_saturation(Some(ImageAdjustment::NEUTRAL))
            .with_enhancement(Some(ImageEnhancement::Disabled));
        let changed =
            __rewrite_image_adjustments_payload(&source, explicit, WireLimits::default()).unwrap();
        assert_eq!(
            __decode_image_adjustments_payload(&changed, WireLimits::default()).unwrap(),
            explicit
        );
        let reset = __rewrite_image_adjustments_payload(
            &changed,
            ImageAdjustments::default(),
            WireLimits::default(),
        )
        .unwrap();
        assert_eq!(reset, source);
    }

    #[test]
    fn focused_bridge_applies_caller_limits() {
        let source = [0x0a, 0x00, 0x72, 0x00];
        let limits = WireLimits::default().with_input_bytes(1).unwrap();
        assert!(__decode_image_adjustments_payload(&source, limits).is_err());
    }
}
