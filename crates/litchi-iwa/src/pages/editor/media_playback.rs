//! Pages-owned MovieArchive playback ingress and mutation.
//!
//! The host owns archive lookup and package transactions.  Scalar playback
//! semantics and wire-preserving rewrites stay in the focused `litchi-pages`
//! package, which receives only one borrowed MovieArchive payload at a time.

use super::*;

use crate::archive::{Archive, ArchiveObject, RawMessage};
use litchi_iwa_common::media::playback::MediaPlaybackSettings;

const MOVIE_ARCHIVE_MESSAGE_TYPE: u32 = 3_007;

pub(super) fn movie_playback_wire_limits(package: &IWorkPackage) -> Result<WireLimits> {
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
            Error::InvalidFormat(format!("invalid Pages movie playback limits: {error}"))
        })
}

pub(super) fn movie_playback_settings(
    package: &IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    context: &str,
) -> Result<MediaPlaybackSettings> {
    let limits = movie_playback_wire_limits(package)?;
    package.with_parsed_archive(archive_name, |archive| {
        let payload = movie_archive_payload(archive, movie_id, context)?;
        litchi_pages::__decode_movie_playback_payload(payload, limits).map_err(|error| {
            Error::InvalidFormat(format!(
                "{context} {movie_id} has invalid playback settings: {error}"
            ))
        })
    })
}

pub(super) fn replace_movie_playback_settings(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    context: &str,
    settings: MediaPlaybackSettings,
) -> Result<MediaPlaybackSettings> {
    let settings = settings
        .canonicalize()
        .map_err(|error| Error::ParseError(error.to_string()))?;
    let limits = movie_playback_wire_limits(package)?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("{context} object {movie_id} is missing"))
        })?;
        let message_index = movie_archive_message_index(object, movie_id, context)?;
        let original = object.messages[message_index].data.as_slice();
        let current =
            litchi_pages::__decode_movie_playback_payload(original, limits).map_err(|error| {
                Error::InvalidFormat(format!(
                    "{context} {movie_id} has invalid playback settings: {error}"
                ))
            })?;
        if current == settings {
            return Ok(());
        }
        let data = litchi_pages::__rewrite_movie_playback_payload(original, settings, limits)
            .map_err(|error| {
                Error::InvalidFormat(format!(
                    "{context} {movie_id} playback rewrite failed: {error}"
                ))
            })?;
        let verified =
            litchi_pages::__decode_movie_playback_payload(&data, limits).map_err(|error| {
                Error::InvalidFormat(format!(
                    "{context} {movie_id} playback rewrite verification failed: {error}"
                ))
            })?;
        if verified != settings {
            return Err(Error::InvalidFormat(format!(
                "{context} playback patch failed validation"
            )));
        }
        object.replace_message(
            message_index,
            RawMessage {
                type_: MOVIE_ARCHIVE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })?;
    Ok(settings)
}

fn movie_archive_payload<'archive>(
    archive: &'archive Archive,
    movie_id: u64,
    context: &str,
) -> Result<&'archive [u8]> {
    let object = archive
        .object(movie_id)
        .ok_or_else(|| Error::InvalidFormat(format!("{context} object {movie_id} is missing")))?;
    let message_index = movie_archive_message_index(object, movie_id, context)?;
    Ok(object.messages[message_index].data.as_slice())
}

fn movie_archive_message_index(
    object: &ArchiveObject,
    movie_id: u64,
    context: &str,
) -> Result<usize> {
    let mut message_index = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != MOVIE_ARCHIVE_MESSAGE_TYPE {
            continue;
        }
        if message_index.replace(index).is_some() {
            return Err(Error::InvalidFormat(format!(
                "{context} {movie_id} must have exactly one MovieArchive payload"
            )));
        }
    }
    message_index.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{context} {movie_id} must have exactly one MovieArchive payload"
        ))
    })
}
