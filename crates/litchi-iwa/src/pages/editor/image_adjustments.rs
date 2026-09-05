//! Pages-owned ImageArchive adjustment ingress and mutation.
//!
//! The host owns archive lookup and package transactions. The focused Pages
//! package owns the bounded semantic projection and source-preserving rewrite
//! of one borrowed ImageArchive payload at a time.

use super::*;

use crate::archive::{ArchiveObject, RawMessage};
use litchi_iwa_common::shape::image::ImageAdjustments;

const IMAGE_ARCHIVE_MESSAGE_TYPE: u32 = 3_005;

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

pub(super) fn replace_image_adjustments(
    package: &mut IWorkPackage,
    archive_name: &str,
    image_id: u64,
    context: &str,
    adjustments: ImageAdjustments,
) -> Result<ImageAdjustments> {
    let limits = image_adjustment_wire_limits(package)?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(image_id).ok_or_else(|| {
            Error::InvalidFormat(format!("{context} object {image_id} is missing"))
        })?;
        let message_index = image_archive_message_index(object, image_id, context)?;
        let original = object.messages[message_index].data.as_slice();
        let current = litchi_pages::__decode_image_adjustments_payload(original, limits).map_err(
            |error| {
                Error::InvalidFormat(format!(
                    "{context} {image_id} has invalid image adjustments: {error}"
                ))
            },
        )?;
        if current == adjustments {
            return Ok(());
        }
        let data = litchi_pages::__rewrite_image_adjustments_payload(original, adjustments, limits)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "{context} {image_id} image adjustment rewrite failed: {error}"
                ))
            })?;
        let verified =
            litchi_pages::__decode_image_adjustments_payload(&data, limits).map_err(|error| {
                Error::InvalidFormat(format!(
                    "{context} {image_id} image adjustment rewrite verification failed: {error}"
                ))
            })?;
        if verified != adjustments {
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

fn image_archive_message_index(
    object: &ArchiveObject,
    image_id: u64,
    context: &str,
) -> Result<usize> {
    let mut message_index = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != IMAGE_ARCHIVE_MESSAGE_TYPE {
            continue;
        }
        if message_index.replace(index).is_some() {
            return Err(Error::InvalidFormat(format!(
                "{context} {image_id} must have exactly one ImageArchive payload"
            )));
        }
    }
    message_index.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{context} {image_id} must have exactly one ImageArchive payload"
        ))
    })
}
