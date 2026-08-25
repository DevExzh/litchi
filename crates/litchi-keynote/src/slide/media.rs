//! Keynote slide media classifications.

use std::time::Duration;

/// A slide-drawable position in document points.
///
/// This is a Keynote-owned semantic value. Native geometry records are decoded
/// into it by the package adapter and never appear in the public slide model.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Point {
    /// Horizontal document coordinate.
    pub x: f32,
    /// Vertical document coordinate.
    pub y: f32,
}

/// A slide-drawable size in document points.
///
/// The dimensions retain the source values, including the distinction between
/// an absent size and a present zero or non-finite native value. Validation of
/// native fields remains in the package adapter.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    /// Width in document points.
    pub width: f32,
    /// Height in document points.
    pub height: f32,
}

/// Validated placement and dimensions for an existing file-backed movie.
///
/// This value is deliberately limited to the semantic geometry that can be
/// shared by Keynote movie readers and writers: a finite top-left position and
/// a finite, strictly positive displayed size. Native flags, angles, archive
/// records, and object identifiers remain in the package adapter.
pub mod geometry {
    use super::{Point, Size};

    /// Semantic validation failures for movie geometry values.
    pub use crate::Error;
    /// Result type for validated movie geometry construction.
    pub type Result<T> = crate::Result<T>;

    /// A validated movie position and displayed size in document points.
    #[derive(Debug, Clone, Copy, PartialEq)]
    pub struct MovieGeometry {
        position: Point,
        size: Size,
    }

    impl MovieGeometry {
        /// Construct validated movie geometry.
        ///
        /// Both position coordinates must be finite. Width and height must
        /// be finite and strictly positive.
        ///
        /// # Errors
        ///
        /// Returns [`crate::Error::InvalidMoviePosition`] for a non-finite
        /// coordinate or [`crate::Error::InvalidMovieSize`] for a non-finite
        /// or non-positive dimension.
        pub const fn new(position: Point, size: Size) -> Result<Self> {
            if !position.x.is_finite() || !position.y.is_finite() {
                return Err(Error::InvalidMoviePosition);
            }
            if !size.width.is_finite()
                || !size.height.is_finite()
                || size.width <= 0.0
                || size.height <= 0.0
            {
                return Err(Error::InvalidMovieSize);
            }
            Ok(Self { position, size })
        }

        /// Return the top-left movie position in document points.
        #[must_use]
        pub const fn position(self) -> Point {
            self.position
        }

        /// Return the displayed movie dimensions in document points.
        #[must_use]
        pub const fn size(self) -> Size {
            self.size
        }
    }

    /// Package transaction types for an existing movie's geometry.
    ///
    /// The package adapter owns selection, native graph traversal, and exact
    /// publication. This namespace exposes only the transaction vocabulary;
    /// native records and identifiers do not cross the semantic boundary.
    pub mod transaction {
        pub use crate::{
            SlideMovieGeometryCommit, SlideMovieGeometryDiagnostics, SlideMovieGeometryEdit,
            SlideMovieGeometryError, SlideMovieGeometryLimitKind, SlideMovieGeometryPatch,
        };
    }
}

/// Archive-free playback values used by Keynote slide media summaries.
pub mod playback {
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

    /// Validation failures for Keynote media playback values.
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
            /// The playback time field that failed validation.
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

    /// A normalized media volume accepted by Keynote.
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

        /// Return the linear volume multiplier.
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

    /// Repeat behavior for Keynote movie and audio clips.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    #[non_exhaustive]
    pub enum MediaLoopMode {
        /// Stop after one playback.
        None,
        /// Restart from the beginning after each playback.
        Repeat,
        /// Alternate forward and reverse playback.
        BackAndForth,
        /// A value introduced by a newer Keynote version.
        Unknown(i32),
    }

    impl MediaLoopMode {
        /// Decode a native movie-loop value losslessly.
        #[must_use]
        pub const fn from_raw(value: i32) -> Self {
            match value {
                NO_LOOP_MODE => Self::None,
                REPEAT_LOOP_MODE => Self::Repeat,
                BACK_AND_FORTH_LOOP_MODE => Self::BackAndForth,
                other => Self::Unknown(other),
            }
        }

        /// Return the native movie-loop value.
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

    /// Playback state retained by a Keynote movie summary.
    ///
    /// `end_time` is required because Keynote uses it as the authoritative
    /// playback boundary. Optional fields preserve the native distinction
    /// between an omitted value and an explicitly encoded default.
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
        /// Validation uses the finite `f32`-seconds representation accepted by
        /// Keynote, so a value that passes this method can cross the package
        /// adapter without a lossy or invalid conversion.
        ///
        /// # Errors
        ///
        /// Returns the same validation error as [`Self::canonicalize`].
        pub fn validate(self) -> Result<()> {
            self.canonicalize().map(|_| ())
        }

        /// Canonicalize durations to Keynote's native `f32`-seconds form.
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
        reason = "Keynote stores media times as f32 seconds"
    )]
    fn canonical_duration(value: Duration, field: TimeField) -> Result<Duration> {
        let seconds = value.as_secs_f64();
        if !seconds.is_finite() || seconds > f64::from(f32::MAX) {
            return Err(Error::TimeOutOfRange { field });
        }
        Duration::try_from_secs_f32(seconds as f32)
            .map_err(|_error| Error::TimeOutOfRange { field })
    }

    /// Package transaction types for an existing file-backed movie's scalar
    /// playback settings.  The transaction itself remains implemented by the
    /// package adapter; this nested namespace keeps the semantic media value
    /// and its operation vocabulary together without exposing native records.
    pub mod transaction {
        pub use crate::{
            SlideMoviePlaybackCommit, SlideMoviePlaybackDiagnostics, SlideMoviePlaybackEdit,
            SlideMoviePlaybackError, SlideMoviePlaybackLimitKind, SlideMoviePlaybackPatch,
        };
    }
}

pub use playback::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};

/// The semantic role of a movie drawable owned directly by a Keynote slide.
///
/// This value deliberately contains no archive, package, or native identifier
/// state. The IWA adapter owns the graph and media records and uses this type
/// only for the product-level classification exposed in movie information.
#[repr(u8)]
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum MovieKind {
    /// An ordinary file-backed movie inserted by the user.
    File,
    /// An independently positioned audio clip stored in a movie archive.
    Audio,
    /// A file-backed replacement target materialized from a slide layout.
    Placeholder,
    /// A camera-backed live-video drawable.
    LiveVideo,
}

impl MovieKind {
    /// Return whether this drawable is an independently positioned audio clip.
    ///
    /// This predicate is future-proof for the non-exhaustive classification:
    /// newly introduced movie kinds remain `false` until they are explicitly
    /// classified as audio.
    #[must_use]
    pub const fn is_audio(self) -> bool {
        matches!(self, Self::Audio)
    }
}

/// An archive-free summary of one movie drawable owned by a slide.
///
/// The summary is intentionally disjoint from the native graph.  It carries
/// source-order media semantics only: no package component names, object
/// identifiers, data-reference identifiers, generated protobuf values, or
/// borrowed package storage cross this boundary.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct MovieInfo {
    kind: MovieKind,
    position: Option<Point>,
    size: Option<Size>,
    original_size: Option<Size>,
    natural_size: Option<Size>,
    playback: Option<MediaPlaybackSettings>,
}

impl MovieInfo {
    /// Construct a detached movie summary from validated semantic values.
    ///
    /// This constructor is primarily an adapter seam: callers provide only
    /// semantic geometry and playback values, never native graph identities.
    /// Values are retained exactly, including absent geometry and playback
    /// fields used by placeholders and live video.
    #[must_use]
    pub const fn from_parts(
        kind: MovieKind,
        position: Option<Point>,
        size: Option<Size>,
        natural_size: Option<Size>,
        playback: Option<MediaPlaybackSettings>,
    ) -> Self {
        Self {
            kind,
            position,
            size,
            original_size: None,
            natural_size,
            playback,
        }
    }

    /// Return a copy carrying the optional native original media dimensions.
    #[must_use]
    pub const fn with_original_size(mut self, original_size: Option<Size>) -> Self {
        self.original_size = original_size;
        self
    }

    /// Return the semantic movie classification.
    #[must_use]
    pub const fn kind(self) -> MovieKind {
        self.kind
    }

    /// Return whether this entry is an independently positioned audio clip.
    #[must_use]
    pub const fn is_audio(self) -> bool {
        self.kind.is_audio()
    }

    /// Return the optional top-left position in document points.
    #[must_use]
    pub const fn position(self) -> Option<Point> {
        self.position
    }

    /// Return the optional displayed dimensions in document points.
    #[must_use]
    pub const fn size(self) -> Option<Size> {
        self.size
    }

    /// Return the optional original media dimensions in document points.
    #[must_use]
    pub const fn original_size(self) -> Option<Size> {
        self.original_size
    }

    /// Return the optional untransformed media dimensions in document points.
    #[must_use]
    pub const fn natural_size(self) -> Option<Size> {
        self.natural_size
    }

    /// Return optional trim, repeat, poster, and volume settings.
    #[must_use]
    pub const fn playback(self) -> Option<MediaPlaybackSettings> {
        self.playback
    }

    /// Return the playable duration when native playback bounds are present.
    #[must_use]
    pub fn duration(self) -> Option<Duration> {
        self.playback.map(MediaPlaybackSettings::duration)
    }
}

/// Compatibility alias emphasizing that a movie entry is a media drawable.
pub type MediaInfo = MovieInfo;

#[cfg(test)]
mod tests {
    use super::{
        MediaInfo, MediaLoopMode, MediaPlaybackSettings, MediaVolume, MovieInfo, MovieKind, Point,
        Size,
    };
    use std::mem::size_of;
    use std::time::Duration;

    #[test]
    fn classification_is_a_compact_copyable_value() {
        assert_eq!(size_of::<MovieKind>(), 1);
        assert_ne!(MovieKind::File, MovieKind::Audio);
        assert_ne!(MovieKind::Placeholder, MovieKind::LiveVideo);
    }

    #[test]
    fn audio_predicate_is_explicit_and_non_exhaustive_safe() {
        assert!(MovieKind::Audio.is_audio());
        assert!(!MovieKind::File.is_audio());
        assert!(!MovieKind::Placeholder.is_audio());
        assert!(!MovieKind::LiveVideo.is_audio());
    }

    #[test]
    fn movie_summary_is_archive_free_and_retains_semantic_presence() {
        let playback = MediaPlaybackSettings::new(Duration::from_secs(3));
        let movie = MovieInfo::from_parts(
            MovieKind::File,
            Some(Point { x: 12.0, y: 24.0 }),
            Some(Size {
                width: 640.0,
                height: 360.0,
            }),
            None,
            Some(playback),
        );

        assert_eq!(movie.kind(), MovieKind::File);
        assert!(!movie.is_audio());
        assert_eq!(movie.position(), Some(Point { x: 12.0, y: 24.0 }));
        assert_eq!(movie.original_size(), None);
        assert_eq!(movie.natural_size(), None);
        assert_eq!(movie.playback(), Some(playback));
        assert_eq!(movie.duration(), Some(Duration::from_secs(3)));
        let _: MediaInfo = movie;
        assert_eq!(size_of::<MovieInfo>(), size_of::<MediaInfo>());
    }

    #[test]
    fn playback_value_keeps_optional_fields_and_unknown_loop_modes() {
        let settings = MediaPlaybackSettings::new(Duration::from_secs(3))
            .with_start_time(Some(Duration::from_secs(1)))
            .with_poster_time(Some(Duration::from_secs(2)))
            .with_loop_mode(Some(MediaLoopMode::Unknown(17)))
            .with_volume(Some(MediaVolume::new(0.75).unwrap()));

        assert_eq!(settings.duration(), Duration::from_secs(2));
        assert!(settings.validate().is_ok());
        assert_eq!(settings.canonicalize().unwrap(), settings);
        assert_eq!(MediaLoopMode::from_raw(17).as_raw(), 17);
        assert_eq!(MediaVolume::FULL.as_f32(), 1.0);
    }
}
