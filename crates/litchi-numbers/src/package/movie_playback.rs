//! Bounded Numbers movie-playback projection and source-preserving rewrite.
//!
//! This module is an internal bridge for the legacy `litchi-iwa` host.  The
//! host borrows one MovieArchive payload from its parsed-archive cache and
//! delegates scalar semantics to the focused Numbers package.  No generated
//! MovieArchive or package-wide source is retained by this seam.

use std::time::Duration;

use litchi_iwa_common::{
    WireLimits,
    media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume, TimeField},
};
use litchi_iwa_protos::movie_playback_codec::{
    self as codec, MoviePlaybackSnapshot, MoviePlaybackWrite,
};
use thiserror::Error;

/// Typed failures at the hidden Numbers MovieArchive playback seam.
#[doc(hidden)]
#[derive(Debug, Clone, PartialEq, Eq, Error)]
pub enum MoviePlaybackError {
    /// The bounded Buffa/wire projection rejected the source or candidate.
    #[error("movie playback codec rejected the payload: {0}")]
    Codec(#[from] codec::DecodeError),
    /// The projected values failed the common media semantic boundary.
    #[error("invalid media playback value: {0}")]
    Semantic(#[from] litchi_iwa_common::media::playback::Error),
    /// A native scalar could not be represented as a Rust duration.
    #[error("media {field:?} time is out of range")]
    TimeOutOfRange {
        /// The scalar that failed conversion.
        field: TimeField,
    },
}

fn codec_options(source: &[u8], limits: WireLimits) -> codec::DecodeOptions {
    let base = codec::DecodeOptions::for_source(source);
    let recursion_limit = u32::try_from(limits.max_nesting())
        .unwrap_or(u32::MAX)
        .min(8);
    base.with_max_input_bytes(base.max_input_bytes().min(limits.max_input_bytes()))
        .with_max_output_bytes(limits.max_output_bytes())
        .with_resource_limits(
            source
                .len()
                .max(1)
                .saturating_mul(4)
                .min(limits.max_fields()),
            source
                .len()
                .max(1)
                .saturating_mul(8)
                .min(limits.max_rewrite_work()),
        )
        .with_recursion_limit(recursion_limit)
}

/// Decode one borrowed Numbers MovieArchive payload into archive-free playback
/// settings.
///
/// This function is hidden because its payload argument is a native wire
/// boundary used only by the umbrella host. Supported Numbers APIs expose
/// playback through semantic movie and audio handles instead.
#[doc(hidden)]
pub fn __movie_playback_settings(
    source: &[u8],
    limits: WireLimits,
) -> Result<MediaPlaybackSettings, MoviePlaybackError> {
    let snapshot = codec::decode_movie_playback(source, codec_options(source, limits))?;
    settings_from_snapshot(snapshot)
}

/// Rewrite one borrowed Numbers MovieArchive payload while preserving unknown
/// fields and their source order.
#[doc(hidden)]
pub fn __rewrite_movie_playback_settings(
    source: &[u8],
    settings: MediaPlaybackSettings,
    limits: WireLimits,
) -> Result<Vec<u8>, MoviePlaybackError> {
    let settings = settings.canonicalize()?;
    let write = write_from_settings(settings)?;
    Ok(codec::rewrite_movie_playback(
        source,
        write,
        codec_options(source, limits),
    )?)
}

fn settings_from_snapshot(
    snapshot: MoviePlaybackSnapshot,
) -> Result<MediaPlaybackSettings, MoviePlaybackError> {
    let start_time = snapshot
        .start_time
        .map(|value| duration_from_seconds(value, TimeField::Start))
        .transpose()?;
    let end_time = duration_from_seconds(snapshot.end_time, TimeField::End)?;
    let poster_time = snapshot
        .poster_time
        .map(|value| duration_from_seconds(value, TimeField::Poster))
        .transpose()?;
    let loop_mode = snapshot.loop_mode.map(MediaLoopMode::from_raw);
    let volume = snapshot.volume.map(MediaVolume::new).transpose()?;
    MediaPlaybackSettings {
        start_time,
        end_time,
        poster_time,
        loop_mode,
        volume,
    }
    .canonicalize()
    .map_err(MoviePlaybackError::from)
}

fn write_from_settings(
    settings: MediaPlaybackSettings,
) -> Result<MoviePlaybackWrite, MoviePlaybackError> {
    Ok(MoviePlaybackWrite::from_values(
        settings
            .start_time
            .map(|value| duration_as_seconds(value, TimeField::Start))
            .transpose()?,
        duration_as_seconds(settings.end_time, TimeField::End)?,
        settings
            .poster_time
            .map(|value| duration_as_seconds(value, TimeField::Poster))
            .transpose()?,
        settings.loop_mode.map(MediaLoopMode::as_raw),
        settings.volume.map(MediaVolume::as_f32),
    ))
}

fn duration_from_seconds(value: f32, field: TimeField) -> Result<Duration, MoviePlaybackError> {
    if !value.is_finite() || value < 0.0 {
        return Err(MoviePlaybackError::TimeOutOfRange { field });
    }
    Duration::try_from_secs_f32(value)
        .map_err(|_error| MoviePlaybackError::TimeOutOfRange { field })
}

fn duration_as_seconds(value: Duration, field: TimeField) -> Result<f32, MoviePlaybackError> {
    let seconds = value.as_secs_f64();
    if !seconds.is_finite() || seconds > f64::from(f32::MAX) {
        return Err(MoviePlaybackError::TimeOutOfRange { field });
    }
    Ok(seconds as f32)
}

#[cfg(test)]
mod tests {
    use std::time::Duration;

    use litchi_iwa_common::WireLimits;
    use litchi_iwa_common::media::playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};

    use super::{__movie_playback_settings, __rewrite_movie_playback_settings};

    fn source() -> Vec<u8> {
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&[0x0a, 0x00]);
        bytes.extend_from_slice(&[0x1d, 0, 0, 0, 0]);
        bytes.extend_from_slice(&[0x25, 0, 0, 0xc0, 0x3f]);
        bytes.extend_from_slice(&[0x2d, 0, 0, 0x80, 0x3f]);
        bytes.extend_from_slice(&[0x3d, 0, 0, 0x80, 0x3f]);
        bytes.extend_from_slice(&[0xc0, 0x01, 0x00]);
        bytes
    }

    #[test]
    fn focused_read_and_rewrite_preserve_semantics_and_unknown_bytes() {
        let mut original = source();
        original.extend_from_slice(&[
            0x98, 0x06, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x00,
        ]);
        let baseline = __movie_playback_settings(&original, WireLimits::default()).unwrap();
        assert_eq!(baseline.end_time, Duration::from_secs_f32(1.5));
        let replacement = MediaPlaybackSettings {
            start_time: Some(Duration::from_millis(250)),
            end_time: Duration::from_millis(1_250),
            poster_time: Some(Duration::from_millis(500)),
            loop_mode: Some(MediaLoopMode::BackAndForth),
            volume: Some(MediaVolume::new(0.75).unwrap()),
        };
        let changed =
            __rewrite_movie_playback_settings(&original, replacement, WireLimits::default())
                .unwrap();
        assert_eq!(
            __movie_playback_settings(&changed, WireLimits::default()).unwrap(),
            replacement.canonicalize().unwrap()
        );
        assert!(changed.ends_with(&[
            0x98, 0x06, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x00
        ]));
        let restored =
            __rewrite_movie_playback_settings(&changed, baseline, WireLimits::default()).unwrap();
        assert_eq!(restored, original);
    }

    #[test]
    fn focused_rewrite_keeps_legacy_only_loop_shape() {
        let mut bytes = vec![0x0a, 0x00, 0x25, 0, 0, 0xc0, 0x3f, 0x30, 0x01];
        let baseline = __movie_playback_settings(&bytes, WireLimits::default()).unwrap();
        assert_eq!(baseline.loop_mode, Some(MediaLoopMode::Repeat));
        let replacement = baseline.with_loop_mode(Some(MediaLoopMode::BackAndForth));
        bytes =
            __rewrite_movie_playback_settings(&bytes, replacement, WireLimits::default()).unwrap();
        assert_eq!(bytes, [0x0a, 0x00, 0x25, 0, 0, 0xc0, 0x3f, 0x30, 0x02]);
        assert!(!bytes.windows(2).any(|window| window == [0xc0, 0x01]));
    }

    #[test]
    fn focused_codec_applies_caller_limits() {
        let bytes = source();
        let limits = WireLimits::default().with_input_bytes(1).unwrap();
        assert!(__movie_playback_settings(&bytes, limits).is_err());
    }
}
