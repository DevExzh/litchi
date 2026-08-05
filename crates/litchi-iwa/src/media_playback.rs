//! Typed, wire-preserving playback settings for iWork movie archives.

use std::time::Duration;

use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tsd;
use crate::wire::{patch_fixed32_field, patch_varint_field};
use crate::{Error, IWorkPackage, Result};

const MOVIE_ARCHIVE_MESSAGE_TYPE: u32 = 3_007;
const START_TIME_FIELD: u32 = 3;
const END_TIME_FIELD: u32 = 4;
const POSTER_TIME_FIELD: u32 = 5;
const LEGACY_LOOP_MODE_FIELD: u32 = 6;
const VOLUME_FIELD: u32 = 7;
const LOOP_MODE_FIELD: u32 = 24;
const NO_LOOP_MODE: i32 = 0;
const REPEAT_LOOP_MODE: i32 = 1;
const BACK_AND_FORTH_LOOP_MODE: i32 = 2;

impl From<litchi_iwa_common::media::playback::Error> for Error {
    fn from(error: litchi_iwa_common::media::playback::Error) -> Self {
        Self::ParseError(error.to_string())
    }
}

pub(crate) fn media_playback_settings(movie: &tsd::MovieArchive) -> Result<MediaPlaybackSettings> {
    let start_time = movie
        .start_time
        .map(|value| duration_from_seconds(value, "media start time"))
        .transpose()?;
    let end_time = movie
        .end_time
        .ok_or_else(|| Error::InvalidFormat("media archive has no end time".to_owned()))
        .and_then(|value| duration_from_seconds(value, "media end time"))?;
    let poster_time = movie
        .poster_time
        .map(|value| duration_from_seconds(value, "media poster time"))
        .transpose()?;
    let loop_mode = movie_loop_mode(movie)?;
    let volume = movie
        .volume
        .map(MediaVolume::new)
        .transpose()
        .map_err(Error::from)?;
    MediaPlaybackSettings {
        start_time,
        end_time,
        poster_time,
        loop_mode,
        volume,
    }
    .canonicalize()
    .map_err(Error::from)
}

pub(crate) fn replace_movie_playback_settings(
    package: &mut IWorkPackage,
    archive_name: &str,
    movie_id: u64,
    context: &str,
    settings: MediaPlaybackSettings,
) -> Result<MediaPlaybackSettings> {
    let settings = settings.canonicalize().map_err(Error::from)?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(movie_id).ok_or_else(|| {
            Error::InvalidFormat(format!("{context} object {movie_id} is missing"))
        })?;
        let message_indexes = object
            .messages
            .iter()
            .enumerate()
            .filter_map(|(index, message)| {
                (message.type_ == MOVIE_ARCHIVE_MESSAGE_TYPE).then_some(index)
            })
            .collect::<Vec<_>>();
        let [message_index] = message_indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "{context} {movie_id} must have exactly one MovieArchive payload"
            )));
        };
        let message_index = *message_index;
        let original = object.messages[message_index].data.as_slice();
        let movie = tsd::MovieArchive::decode(original)?;
        if media_playback_settings(&movie).is_ok_and(|current| current == settings) {
            return Ok(());
        }
        let data = patch_movie_playback_settings(original, &movie, settings)?;
        let verified = tsd::MovieArchive::decode(data.as_slice())?;
        if media_playback_settings(&verified)? != settings {
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

#[allow(deprecated)]
fn patch_movie_playback_settings(
    original: &[u8],
    movie: &tsd::MovieArchive,
    settings: MediaPlaybackSettings,
) -> Result<Vec<u8>> {
    let data = patch_fixed32_field(
        original,
        START_TIME_FIELD,
        movie.start_time.is_some(),
        settings
            .start_time
            .map(|value| duration_as_seconds(value, "media start time"))
            .transpose()?
            .map(f32::to_bits),
    )?;
    let data = patch_fixed32_field(
        &data,
        END_TIME_FIELD,
        movie.end_time.is_some(),
        Some(duration_as_seconds(settings.end_time, "media end time")?.to_bits()),
    )?;
    let data = patch_fixed32_field(
        &data,
        POSTER_TIME_FIELD,
        movie.poster_time.is_some(),
        settings
            .poster_time
            .map(|value| duration_as_seconds(value, "media poster time"))
            .transpose()?
            .map(f32::to_bits),
    )?;
    let data = patch_varint_field(
        &data,
        LEGACY_LOOP_MODE_FIELD,
        movie.loop_option_as_integer.is_some(),
        if movie.loop_option_as_integer.is_some() {
            legacy_loop_mode_value(settings.loop_mode)?
        } else {
            None
        },
    )?;
    let data = patch_varint_field(
        &data,
        LOOP_MODE_FIELD,
        movie.loop_option.is_some(),
        if movie.loop_option.is_some() || movie.loop_option_as_integer.is_none() {
            settings
                .loop_mode
                .map(|mode| i64::from(mode.as_raw()) as u64)
        } else {
            None
        },
    )?;
    patch_fixed32_field(
        &data,
        VOLUME_FIELD,
        movie.volume.is_some(),
        settings.volume.map(|volume| volume.as_f32().to_bits()),
    )
}

#[allow(deprecated)]
fn movie_loop_mode(movie: &tsd::MovieArchive) -> Result<Option<MediaLoopMode>> {
    let modern = movie.loop_option.map(MediaLoopMode::from_raw);
    let legacy = movie
        .loop_option_as_integer
        .map(|value| MediaLoopMode::from_raw(i32::from_le_bytes(value.to_le_bytes())));
    match (modern, legacy) {
        (Some(modern), Some(legacy)) if modern != legacy => Err(Error::InvalidFormat(
            "media archive has conflicting modern and legacy loop modes".to_owned(),
        )),
        (Some(mode), _) | (None, Some(mode)) => Ok(Some(mode)),
        (None, None) => Ok(None),
    }
}

fn legacy_loop_mode_value(loop_mode: Option<MediaLoopMode>) -> Result<Option<u64>> {
    Ok(loop_mode.map(|mode| u64::from(u32::from_le_bytes(mode.as_raw().to_le_bytes()))))
}

fn canonical_duration(value: Duration, context: &str) -> Result<Duration> {
    duration_from_seconds(duration_as_seconds(value, context)?, context)
}

fn duration_as_seconds(value: Duration, context: &str) -> Result<f32> {
    let seconds = value.as_secs_f64();
    if !seconds.is_finite() || seconds > f64::from(f32::MAX) {
        return Err(Error::ParseError(format!(
            "{context} must fit in finite f32 seconds"
        )));
    }
    Ok(seconds as f32)
}

fn duration_from_seconds(value: f32, context: &str) -> Result<Duration> {
    if !value.is_finite() || value < 0.0 {
        return Err(Error::InvalidFormat(format!(
            "{context} must be finite and non-negative"
        )));
    }
    Duration::try_from_secs_f32(value)
        .map_err(|error| Error::InvalidFormat(format!("{context} is out of range: {error}")))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn playback_patch_preserves_unknown_fields_and_restores_exactly() {
        let movie = tsd::MovieArchive {
            super_: tsd::DrawableArchive::default(),
            start_time: Some(0.0),
            end_time: Some(1.5),
            poster_time: Some(0.0),
            loop_option: Some(NO_LOOP_MODE),
            volume: Some(1.0),
            ..Default::default()
        };
        let baseline = media_playback_settings(&movie).unwrap();
        let replacement = MediaPlaybackSettings {
            start_time: Some(Duration::from_millis(250)),
            end_time: Duration::from_millis(1_250),
            poster_time: Some(Duration::from_millis(500)),
            loop_mode: Some(MediaLoopMode::BackAndForth),
            volume: Some(MediaVolume::new(0.75).unwrap()),
        };
        let mut original = movie.encode_to_vec();
        append_unknown_varint(&mut original, 99, 990);

        let changed = patch_movie_playback_settings(&original, &movie, replacement).unwrap();
        let changed_movie = tsd::MovieArchive::decode(changed.as_slice()).unwrap();
        assert_eq!(
            media_playback_settings(&changed_movie).unwrap(),
            replacement
        );
        assert!(
            changed
                .windows(3)
                .any(|window| window == [0x98, 0x06, 0xde])
        );

        let restored = patch_movie_playback_settings(&changed, &changed_movie, baseline).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    #[allow(deprecated)]
    fn playback_settings_read_legacy_loop_modes() {
        let movie = tsd::MovieArchive {
            super_: tsd::DrawableArchive::default(),
            end_time: Some(1.0),
            loop_option_as_integer: Some(REPEAT_LOOP_MODE as u32),
            ..Default::default()
        };
        assert_eq!(
            media_playback_settings(&movie).unwrap().loop_mode,
            Some(MediaLoopMode::Repeat)
        );
    }

    #[test]
    #[allow(deprecated)]
    fn playback_settings_round_trip_signed_unknown_legacy_loop_modes() {
        let movie = tsd::MovieArchive {
            super_: tsd::DrawableArchive::default(),
            end_time: Some(1.0),
            loop_option_as_integer: Some(u32::from_le_bytes((-1_i32).to_le_bytes())),
            ..Default::default()
        };
        let settings = media_playback_settings(&movie).unwrap();
        assert_eq!(settings.loop_mode, Some(MediaLoopMode::Unknown(-1)));

        let patched =
            patch_movie_playback_settings(&movie.encode_to_vec(), &movie, settings).unwrap();
        let restored = tsd::MovieArchive::decode(patched.as_slice()).unwrap();
        assert_eq!(
            restored.loop_option_as_integer,
            movie.loop_option_as_integer
        );
    }

    #[test]
    #[allow(deprecated)]
    fn legacy_only_loop_fields_are_not_upgraded() {
        let movie = tsd::MovieArchive {
            super_: tsd::DrawableArchive::default(),
            end_time: Some(1.5),
            loop_option_as_integer: Some(REPEAT_LOOP_MODE as u32),
            ..Default::default()
        };
        let baseline = media_playback_settings(&movie).unwrap();
        let replacement = MediaPlaybackSettings {
            loop_mode: Some(MediaLoopMode::BackAndForth),
            ..baseline
        };
        let original = movie.encode_to_vec();

        let changed = patch_movie_playback_settings(&original, &movie, replacement).unwrap();
        let changed_movie = tsd::MovieArchive::decode(changed.as_slice()).unwrap();
        assert_eq!(changed_movie.loop_option, None);
        assert_eq!(
            changed_movie.loop_option_as_integer,
            Some(BACK_AND_FORTH_LOOP_MODE as u32)
        );

        let restored = patch_movie_playback_settings(&changed, &changed_movie, baseline).unwrap();
        assert_eq!(restored, original);
    }

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(field_number) << 3,
        ));
        data.extend(litchi_iwa_common::varint::encode_varint(value));
    }
}
