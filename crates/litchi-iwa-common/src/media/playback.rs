//! Archive-free semantic playback values for iWork media.
//!
//! Movie-archive decoding, legacy-field reconciliation, wire-preserving
//! patches, and package transactions stay in `litchi-iwa`. This module owns
//! only the compact values exchanged at that boundary and their validation.

use std::time::Duration;

const NO_LOOP_MODE: i32 = 0;
const REPEAT_LOOP_MODE: i32 = 1;
const BACK_AND_FORTH_LOOP_MODE: i32 = 2;

/// The time field that failed native `f32`-second validation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TimeField {
    /// The optional trim start.
    Start,
    /// The required trim end.
    End,
    /// The optional poster position.
    Poster,
}

/// Validation failures for media playback values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, thiserror::Error)]
pub enum Error {
    /// The supplied volume was NaN or infinite.
    #[error("media volume must be finite")]
    VolumeNonFinite,
    /// The supplied volume was outside the native inclusive range.
    #[error("media volume must be in 0.0..=1.0")]
    VolumeOutOfRange,
    /// A duration cannot be represented as finite native `f32` seconds.
    #[error("media {field:?} time must fit in finite f32 seconds")]
    TimeOutOfRange {
        /// The duration field that failed validation.
        field: TimeField,
    },
    /// The trim end was not later than the effective trim start.
    #[error("media end time must be later than its start time")]
    EndTimeNotAfterStart,
    /// A known native loop value was incorrectly wrapped as `Unknown`.
    #[error("media loop mode must not use a reserved native value as unknown")]
    NonCanonicalLoopMode,
}

/// Result type for media playback value construction and validation.
pub type Result<T> = std::result::Result<T, Error>;

/// A normalized media volume accepted by Pages, Numbers, and Keynote.
///
/// Values are expressed as a linear multiplier in the inclusive range
/// `0.0..=1.0`. Construction rejects non-finite and out-of-range values.
#[derive(Debug, Clone, Copy, PartialEq)]
#[repr(transparent)]
pub struct MediaVolume(f32);

impl MediaVolume {
    /// Silence the media clip.
    pub const SILENT: Self = Self(0.0);
    /// Play the media clip at its unattenuated source volume.
    pub const FULL: Self = Self(1.0);

    /// Construct one validated linear volume multiplier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::VolumeNonFinite`] for NaN or infinity and
    /// [`Error::VolumeOutOfRange`] outside the inclusive native range.
    #[must_use = "use the validated volume or handle its validation error"]
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::VolumeNonFinite);
        }
        if !(0.0..=1.0).contains(&value) {
            return Err(Error::VolumeOutOfRange);
        }
        Ok(Self(value))
    }

    /// Return the native linear volume multiplier.
    #[must_use]
    pub const fn as_f32(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for MediaVolume {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::new(value)
    }
}

/// Repeat behavior shared by movie and audio clips.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum MediaLoopMode {
    /// Stop after one playback.
    None,
    /// Restart from the beginning after each playback.
    Repeat,
    /// Alternate forward and reverse playback.
    BackAndForth,
    /// A value introduced by a newer iWork version.
    Unknown(i32),
}

impl MediaLoopMode {
    /// Decode a native `TSD.MovieArchive.MovieLoopOption` value losslessly.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            NO_LOOP_MODE => Self::None,
            REPEAT_LOOP_MODE => Self::Repeat,
            BACK_AND_FORTH_LOOP_MODE => Self::BackAndForth,
            other => Self::Unknown(other),
        }
    }

    /// Return the native `TSD.MovieArchive.MovieLoopOption` value.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::None => NO_LOOP_MODE,
            Self::Repeat => REPEAT_LOOP_MODE,
            Self::BackAndForth => BACK_AND_FORTH_LOOP_MODE,
            Self::Unknown(value) => value,
        }
    }

    /// Return whether this value uses a named variant for a known value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(NO_LOOP_MODE | REPEAT_LOOP_MODE | BACK_AND_FORTH_LOOP_MODE)
        )
    }
}

/// Playback state stored by an iWork movie archive.
///
/// `end_time` is required because the supported file-backed media graphs use
/// it as the authoritative playback boundary. Optional fields preserve the
/// native distinction between an omitted value and an explicitly encoded
/// default.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MediaPlaybackSettings {
    /// Optional absolute trim start from the beginning of the source media.
    pub start_time: Option<Duration>,
    /// Absolute trim end from the beginning of the source media.
    pub end_time: Duration,
    /// Optional absolute frame or sample position used for the media poster.
    pub poster_time: Option<Duration>,
    /// Optional repeat behavior.
    pub loop_mode: Option<MediaLoopMode>,
    /// Optional linear volume multiplier.
    pub volume: Option<MediaVolume>,
}

impl MediaPlaybackSettings {
    /// Create settings with an explicit playback end and no optional fields.
    #[must_use]
    pub const fn new(end_time: Duration) -> Self {
        Self {
            start_time: None,
            end_time,
            poster_time: None,
            loop_mode: None,
            volume: None,
        }
    }

    /// Set the optional absolute trim start.
    #[must_use]
    pub const fn with_start_time(mut self, start_time: Option<Duration>) -> Self {
        self.start_time = start_time;
        self
    }

    /// Set the optional absolute poster position.
    #[must_use]
    pub const fn with_poster_time(mut self, poster_time: Option<Duration>) -> Self {
        self.poster_time = poster_time;
        self
    }

    /// Set the optional repeat behavior.
    #[must_use]
    pub const fn with_loop_mode(mut self, loop_mode: Option<MediaLoopMode>) -> Self {
        self.loop_mode = loop_mode;
        self
    }

    /// Set the optional linear volume multiplier.
    #[must_use]
    pub const fn with_volume(mut self, volume: Option<MediaVolume>) -> Self {
        self.volume = volume;
        self
    }

    /// Validate these settings without changing them.
    ///
    /// The native movie archive stores durations as `f32` seconds. Validation
    /// therefore uses the same canonicalization as the IWA adapter, ensuring
    /// a value that passes this method can be published without a lossy or
    /// invalid conversion.
    ///
    /// # Errors
    ///
    /// Returns the same validation error as [`Self::canonicalize`].
    pub fn validate(self) -> Result<()> {
        self.canonicalize().map(|_| ())
    }

    /// Canonicalize durations to the native `f32`-seconds representation.
    ///
    /// This is intentionally fallible: the IWA adapter uses the returned
    /// value for post-patch comparison, so subrepresentable or invalid trim
    /// ranges cannot silently cross the semantic boundary.
    ///
    /// # Errors
    ///
    /// Returns [`Error::TimeOutOfRange`] for a duration that cannot be
    /// represented as native `f32` seconds, [`Error::EndTimeNotAfterStart`]
    /// for an empty or reversed trim range, and
    /// [`Error::NonCanonicalLoopMode`] for a known value wrapped as unknown.
    pub fn canonicalize(self) -> Result<Self> {
        let start_time = self
            .start_time
            .map(|value| canonical_duration(value, TimeField::Start))
            .transpose()?;
        let end_time = canonical_duration(self.end_time, TimeField::End)?;
        let poster_time = self
            .poster_time
            .map(|value| canonical_duration(value, TimeField::Poster))
            .transpose()?;
        let start = start_time.unwrap_or(Duration::ZERO);
        if end_time <= start {
            return Err(Error::EndTimeNotAfterStart);
        }
        if let Some(loop_mode) = self.loop_mode
            && !loop_mode.is_canonical()
        {
            return Err(Error::NonCanonicalLoopMode);
        }
        Ok(Self {
            start_time,
            end_time,
            poster_time,
            loop_mode: self.loop_mode,
            volume: self.volume,
        })
    }

    /// Return the playable duration after applying the trim range.
    #[must_use]
    pub fn duration(self) -> Duration {
        self.end_time
            .saturating_sub(self.start_time.unwrap_or(Duration::ZERO))
    }
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "the native movie schema is explicitly f32 seconds"
)]
fn canonical_duration(value: Duration, field: TimeField) -> Result<Duration> {
    let seconds = value.as_secs_f64();
    if !seconds.is_finite() || seconds > f64::from(f32::MAX) {
        return Err(Error::TimeOutOfRange { field });
    }
    Duration::try_from_secs_f32(seconds as f32).map_err(|_error| Error::TimeOutOfRange { field })
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;
    use std::time::Duration;

    use super::{Error, MediaLoopMode, MediaPlaybackSettings, MediaVolume};

    #[test]
    fn volume_is_compact_and_strictly_validated() {
        assert_eq!(size_of::<MediaVolume>(), size_of::<f32>());
        assert_eq!(MediaVolume::new(f32::NAN), Err(Error::VolumeNonFinite));
        assert_eq!(MediaVolume::new(f32::INFINITY), Err(Error::VolumeNonFinite));
        assert_eq!(
            MediaVolume::new(-f32::EPSILON),
            Err(Error::VolumeOutOfRange)
        );
        assert_eq!(
            MediaVolume::new(1.0 + f32::EPSILON),
            Err(Error::VolumeOutOfRange)
        );
        assert_eq!(MediaVolume::SILENT.as_f32(), 0.0);
        assert_eq!(MediaVolume::FULL.as_f32(), 1.0);
    }

    #[test]
    fn loop_modes_round_trip_unknown_values_without_shadowing_known_modes() {
        for raw in [i32::MIN, -1, 0, 1, 2, 17, i32::MAX] {
            assert_eq!(MediaLoopMode::from_raw(raw).as_raw(), raw);
        }
        assert_eq!(MediaLoopMode::from_raw(0), MediaLoopMode::None);
        assert_eq!(MediaLoopMode::from_raw(1), MediaLoopMode::Repeat);
        assert_eq!(MediaLoopMode::from_raw(2), MediaLoopMode::BackAndForth);
        for raw in [0, 1, 2] {
            assert!(!MediaLoopMode::Unknown(raw).is_canonical());
        }
        assert!(MediaLoopMode::Unknown(-1).is_canonical());
        assert!(MediaLoopMode::Unknown(17).is_canonical());
    }

    #[test]
    fn settings_builders_and_validation_preserve_optional_presence() {
        let settings = MediaPlaybackSettings::new(Duration::from_secs(3))
            .with_start_time(Some(Duration::from_secs(1)))
            .with_poster_time(Some(Duration::from_secs(2)))
            .with_loop_mode(Some(MediaLoopMode::Repeat))
            .with_volume(Some(MediaVolume::new(0.75).unwrap()));
        assert_eq!(settings.duration(), Duration::from_secs(2));
        assert!(settings.validate().is_ok());
        assert_eq!(settings.canonicalize().unwrap(), settings);
    }

    #[test]
    fn settings_reject_invalid_ranges_and_reserved_unknown_modes() {
        assert_eq!(
            MediaPlaybackSettings::new(Duration::ZERO).validate(),
            Err(Error::EndTimeNotAfterStart)
        );
        assert_eq!(
            MediaPlaybackSettings::new(Duration::from_secs(1))
                .with_start_time(Some(Duration::from_secs(1)))
                .validate(),
            Err(Error::EndTimeNotAfterStart)
        );
        assert_eq!(
            MediaPlaybackSettings::new(Duration::from_secs(1))
                .with_loop_mode(Some(MediaLoopMode::Unknown(0)))
                .validate(),
            Err(Error::NonCanonicalLoopMode)
        );
    }
}
