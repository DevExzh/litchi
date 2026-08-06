//! Archive-free Keynote soundtrack playback values.

const PLAY_ONCE_MODE: i32 = 0;
const LOOP_MODE: i32 = 1;
const DO_NOT_PLAY_MODE: i32 = 2;

/// How Keynote plays a presentation soundtrack.
///
/// Unknown native discriminants are retained so reading a newer Keynote file
/// does not discard information. [`Settings::validate`] rejects spelling a
/// known value as `Unknown`, while accepting genuinely future values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Mode {
    /// Play the soundtrack once.
    PlayOnce,
    /// Repeat the soundtrack.
    Loop,
    /// Do not play the soundtrack.
    DoNotPlay,
    /// A mode introduced by a newer Keynote release.
    Unknown(i32),
}

impl Mode {
    /// Decode a native `KN.Soundtrack.SoundtrackMode` discriminant losslessly.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            PLAY_ONCE_MODE => Self::PlayOnce,
            LOOP_MODE => Self::Loop,
            DO_NOT_PLAY_MODE => Self::DoNotPlay,
            other => Self::Unknown(other),
        }
    }

    /// Return the native `KN.Soundtrack.SoundtrackMode` discriminant.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::PlayOnce => PLAY_ONCE_MODE,
            Self::Loop => LOOP_MODE,
            Self::DoNotPlay => DO_NOT_PLAY_MODE,
            Self::Unknown(value) => value,
        }
    }

    /// Return whether this value uses a named variant for a known native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(PLAY_ONCE_MODE | LOOP_MODE | DO_NOT_PLAY_MODE)
        )
    }
}

/// Validated playback settings for a presentation soundtrack.
///
/// Media entries are deliberately absent. They are package-owned resources
/// with native data-reference metadata and are edited through the IWA
/// soundtrack-item API. Settings edits therefore preserve that collection
/// without exposing archive topology in this semantic value.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct Settings {
    /// Native playback volume in the inclusive range `0.0..=1.0`.
    volume: Option<f64>,
    /// Optional native playback mode.
    mode: Option<Mode>,
}

impl Settings {
    /// Construct playback settings from optional native values.
    ///
    /// Values are checked before the settings are returned, so a safe caller
    /// cannot publish an out-of-range volume or a known mode disguised as an
    /// unknown value through this semantic type.
    ///
    /// # Errors
    ///
    /// Returns a typed error when either optional value is invalid.
    pub const fn new(volume: Option<f64>, mode: Option<Mode>) -> Result<Self, Error> {
        let settings = Self { volume, mode };
        match settings.validate() {
            Ok(()) => Ok(settings),
            Err(error) => Err(error),
        }
    }

    /// Return the optional native playback volume.
    #[must_use]
    pub const fn volume(self) -> Option<f64> {
        self.volume
    }

    /// Replace or clear the native playback volume after validating it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFiniteVolume`] or [`Error::VolumeOutOfRange`] when
    /// `volume` is not representable by the native field.
    pub const fn set_volume(&mut self, volume: Option<f64>) -> Result<(), Error> {
        if let Err(error) = validate_volume(volume) {
            return Err(error);
        }
        self.volume = volume;
        Ok(())
    }

    /// Return the optional native playback mode.
    #[must_use]
    pub const fn mode(self) -> Option<Mode> {
        self.mode
    }

    /// Replace or clear the native playback mode after validating it.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalMode`] when a known native discriminant
    /// is passed through [`Mode::Unknown`].
    pub const fn set_mode(&mut self, mode: Option<Mode>) -> Result<(), Error> {
        if let Some(candidate_mode) = mode
            && !candidate_mode.is_canonical()
        {
            return Err(Error::NonCanonicalMode);
        }
        self.mode = mode;
        Ok(())
    }

    /// Validate values before they cross into a package adapter.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFiniteVolume`] or [`Error::VolumeOutOfRange`] for
    /// an invalid volume, and [`Error::NonCanonicalMode`] when a known native
    /// discriminant is wrapped in [`Mode::Unknown`].
    pub const fn validate(self) -> Result<(), Error> {
        if let Err(error) = validate_volume(self.volume) {
            return Err(error);
        }
        if let Some(mode) = self.mode
            && !mode.is_canonical()
        {
            return Err(Error::NonCanonicalMode);
        }
        Ok(())
    }
}

/// A soundtrack semantic value failed validation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// Volume was NaN or infinite.
    NonFiniteVolume,
    /// Volume was outside the native inclusive range.
    VolumeOutOfRange,
    /// A known native discriminant was wrapped as `Mode::Unknown`.
    NonCanonicalMode,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::NonFiniteVolume => "soundtrack volume must be finite",
            Self::VolumeOutOfRange => "soundtrack volume must be between zero and one",
            Self::NonCanonicalMode => {
                "soundtrack mode must use its named variant for known native values"
            },
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

const fn validate_volume(candidate: Option<f64>) -> Result<(), Error> {
    let Some(volume) = candidate else {
        return Ok(());
    };
    if !volume.is_finite() {
        return Err(Error::NonFiniteVolume);
    }
    if volume < 0.0 || volume > 1.0 {
        return Err(Error::VolumeOutOfRange);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn modes_map_native_values_losslessly() {
        for raw in [0, 1, 2, 19, -1, i32::MIN, i32::MAX] {
            assert_eq!(Mode::from_raw(raw).as_raw(), raw);
        }
        assert_eq!(Mode::from_raw(0), Mode::PlayOnce);
        assert_eq!(Mode::from_raw(1), Mode::Loop);
        assert_eq!(Mode::from_raw(2), Mode::DoNotPlay);
        assert_eq!(Mode::from_raw(19), Mode::Unknown(19));
    }

    #[test]
    fn known_values_cannot_be_smuggled_as_unknown() {
        for raw in [0, 1, 2] {
            assert!(!Mode::Unknown(raw).is_canonical());
        }
        assert!(Mode::Unknown(-1).is_canonical());
        assert!(Mode::Unknown(i32::MAX).is_canonical());
    }

    #[test]
    fn settings_validate_volume_boundaries_and_modes() {
        for volume in [None, Some(0.0), Some(1.0)] {
            assert!(Settings::new(volume, Some(Mode::PlayOnce)).is_ok());
        }
        for volume in [
            Some(-f64::EPSILON),
            Some(1.0 + f64::EPSILON),
            Some(f64::NAN),
            Some(f64::INFINITY),
            Some(f64::NEG_INFINITY),
        ] {
            assert!(Settings::new(volume, None).is_err());
        }
        assert!(Settings::new(None, Some(Mode::Unknown(19))).is_ok());
        assert_eq!(
            Settings::new(None, Some(Mode::Unknown(1))).map(|_| ()),
            Err(Error::NonCanonicalMode)
        );
    }

    #[test]
    fn setters_validate_before_mutating() {
        let mut settings = Settings::default();
        assert_eq!(settings.set_volume(Some(0.5)), Ok(()));
        assert_eq!(settings.volume(), Some(0.5));
        for volume in [
            Some(-f64::EPSILON),
            Some(1.0 + f64::EPSILON),
            Some(f64::NAN),
            Some(f64::INFINITY),
            Some(f64::NEG_INFINITY),
        ] {
            assert!(settings.set_volume(volume).is_err());
            assert_eq!(settings.volume(), Some(0.5));
        }
        assert_eq!(settings.set_mode(Some(Mode::Loop)), Ok(()));
        assert_eq!(settings.mode(), Some(Mode::Loop));
        assert_eq!(
            settings.set_mode(Some(Mode::Unknown(2))),
            Err(Error::NonCanonicalMode)
        );
        assert_eq!(settings.mode(), Some(Mode::Loop));
    }
}
