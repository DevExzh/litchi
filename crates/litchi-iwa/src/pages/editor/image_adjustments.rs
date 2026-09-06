//! Pages-owned ImageArchive adjustment ingress.
//!
//! The focused Pages package owns body-image adjustment mutation. The host
//! retains only the bounded semantic projection needed by the image-info
//! graph reader.

use super::*;

use litchi_iwa_common::shape::image::ImageAdjustments;

pub(super) fn image_adjustment_wire_limits(package: &IWorkPackage) -> Result<WireLimits> {
    let archive_limits = package.limits().archive_limits();
    let source_bytes = archive_limits
        .max_message_bytes()
        .min(archive_limits.max_archive_bytes())
        .min(package.limits().max_iwa_stream_bytes())
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    WireLimits::default()
        .with_input_bytes(source_bytes)
        .and_then(|limits| {
            limits.with_fields(
                source_bytes
                    .saturating_mul(4)
                    .clamp(1, WireLimits::MAX_FIELDS),
            )
        })
        .and_then(|limits| limits.with_output_bytes(source_bytes))
        .and_then(|limits| {
            limits.with_rewrite_work(
                source_bytes
                    .saturating_mul(8)
                    .clamp(1, WireLimits::MAX_REWRITE_WORK),
            )
        })
        .map_err(|error| {
            Error::InvalidFormat(format!("invalid Pages image adjustment limits: {error}"))
        })
}

pub(super) fn image_adjustments_from_payload(
    package: &IWorkPackage,
    payload: &[u8],
    image_id: u64,
    context: &str,
) -> Result<ImageAdjustments> {
    let limits = image_adjustment_wire_limits(package)?;
    litchi_pages::__decode_image_adjustments_payload(payload, limits).map_err(|error| {
        Error::InvalidFormat(format!(
            "{context} {image_id} has invalid image adjustments: {error}"
        ))
    })
}
