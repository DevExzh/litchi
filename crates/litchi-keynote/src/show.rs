//! Presentation-level Keynote show values.

use core::fmt;

const NORMAL_MODE: i32 = 0;
const SELF_PLAYING_MODE: i32 = 1;
const LINKS_ONLY_MODE: i32 = 2;

/// A finite, positive presentation dimension pair in points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    width: f32,
    height: f32,
}

impl Size {
    /// Construct checked presentation dimensions.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDimensions`] when either dimension is not
    /// finite and strictly positive.
    pub fn new(width: f32, height: f32) -> Result<Self, Error> {
        if width.is_finite() && width > 0.0 && height.is_finite() && height > 0.0 {
            Ok(Self { width, height })
        } else {
            Err(Error::InvalidDimensions)
        }
    }

    /// Return the presentation width in points.
    #[must_use]
    pub const fn width(self) -> f32 {
        self.width
    }

    /// Return the presentation height in points.
    #[must_use]
    pub const fn height(self) -> f32 {
        self.height
    }
}

/// Errors returned while constructing or validating presentation settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Error {
    /// A width or height was not finite and strictly positive.
    InvalidDimensions,
    /// A playback delay was not finite and non-negative.
    InvalidDelay,
    /// A known native mode was supplied as an `Unknown` value.
    NonCanonicalMode,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let message = match self {
            Self::InvalidDimensions => "show dimensions must be finite and greater than zero",
            Self::InvalidDelay => "show playback delays must be finite and non-negative",
            Self::NonCanonicalMode => "show mode must use its named variant for known values",
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// How a Keynote presentation advances and responds to input.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Mode {
    /// Slides advance normally through presenter input and configured transitions.
    Normal,
    /// The presentation advances automatically using its playback delays.
    SelfPlaying,
    /// Only hyperlinks can navigate between slides.
    LinksOnly,
    /// A mode introduced by a newer Keynote version.
    Unknown(i32),
}

impl Mode {
    /// Decode a native `KNShowMode` value without discarding unknown values.
    #[must_use]
    pub const fn from_raw(value: i32) -> Self {
        match value {
            NORMAL_MODE => Self::Normal,
            SELF_PLAYING_MODE => Self::SelfPlaying,
            LINKS_ONLY_MODE => Self::LinksOnly,
            raw => Self::Unknown(raw),
        }
    }

    /// Construct an unknown mode, rejecting values already assigned to a named mode.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalMode`] when `value` is already assigned to
    /// one of the named mode variants.
    pub const fn unknown(value: i32) -> Result<Self, Error> {
        match value {
            NORMAL_MODE | SELF_PLAYING_MODE | LINKS_ONLY_MODE => Err(Error::NonCanonicalMode),
            raw => Ok(Self::Unknown(raw)),
        }
    }

    /// Return the native `KNShowMode` value stored in the Keynote archive.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Normal => NORMAL_MODE,
            Self::SelfPlaying => SELF_PLAYING_MODE,
            Self::LinksOnly => LINKS_ONLY_MODE,
            Self::Unknown(value) => value,
        }
    }

    /// Return whether this value is the canonical representation of its native value.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(NORMAL_MODE | SELF_PLAYING_MODE | LINKS_ONLY_MODE)
        )
    }
}

/// Validated presentation dimensions and playback behavior.
#[derive(Debug, Clone, PartialEq)]
pub struct Settings {
    size: Size,
    slide_numbers_visible: Option<bool>,
    loop_presentation: Option<bool>,
    mode: Option<Mode>,
    autoplay_transition_delay: Option<f64>,
    autoplay_build_delay: Option<f64>,
    idle_timer_active: Option<bool>,
    idle_timer_delay: Option<f64>,
    automatically_plays_upon_open: Option<bool>,
}

impl Settings {
    /// Construct default playback settings for checked presentation dimensions.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDimensions`] when either dimension is not
    /// finite and strictly positive.
    pub fn new(width: f32, height: f32) -> Result<Self, Error> {
        Ok(Self {
            size: Size::new(width, height)?,
            slide_numbers_visible: None,
            loop_presentation: None,
            mode: None,
            autoplay_transition_delay: None,
            autoplay_build_delay: None,
            idle_timer_active: None,
            idle_timer_delay: None,
            automatically_plays_upon_open: None,
        })
    }

    /// Return the checked presentation dimensions.
    #[must_use]
    pub const fn size(&self) -> Size {
        self.size
    }

    /// Replace the presentation dimensions after validation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDimensions`] when either dimension is not
    /// finite and strictly positive. The previous size is retained on error.
    pub fn set_size(&mut self, width: f32, height: f32) -> Result<(), Error> {
        self.size = Size::new(width, height)?;
        Ok(())
    }

    /// Return whether slide numbers are visible when the setting is present.
    #[must_use]
    pub const fn slide_numbers_visible(&self) -> Option<bool> {
        self.slide_numbers_visible
    }

    /// Set or clear the slide-number visibility override.
    pub const fn set_slide_numbers_visible(&mut self, value: Option<bool>) {
        self.slide_numbers_visible = value;
    }

    /// Return whether the presentation loops when the setting is present.
    #[must_use]
    pub const fn loop_presentation(&self) -> Option<bool> {
        self.loop_presentation
    }

    /// Set or clear the loop override.
    pub const fn set_loop_presentation(&mut self, value: Option<bool>) {
        self.loop_presentation = value;
    }

    /// Return the playback mode when explicitly stored.
    #[must_use]
    pub const fn mode(&self) -> Option<Mode> {
        self.mode
    }

    /// Set or clear the playback mode after canonical-value validation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalMode`] for an `Unknown` value whose raw
    /// value is already assigned to a named mode. The previous mode is
    /// retained on error.
    pub const fn set_mode(&mut self, value: Option<Mode>) -> Result<(), Error> {
        if let Some(mode) = value
            && !mode.is_canonical()
        {
            return Err(Error::NonCanonicalMode);
        }
        self.mode = value;
        Ok(())
    }

    /// Return the automatic transition delay in seconds.
    #[must_use]
    pub const fn autoplay_transition_delay(&self) -> Option<f64> {
        self.autoplay_transition_delay
    }

    /// Set or clear the automatic transition delay.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDelay`] when the value is present but not
    /// finite and non-negative. The previous delay is retained on error.
    pub fn set_autoplay_transition_delay(&mut self, value: Option<f64>) -> Result<(), Error> {
        validate_delay(value)?;
        self.autoplay_transition_delay = value;
        Ok(())
    }

    /// Return the automatic build delay in seconds.
    #[must_use]
    pub const fn autoplay_build_delay(&self) -> Option<f64> {
        self.autoplay_build_delay
    }

    /// Set or clear the automatic build delay.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDelay`] when the value is present but not
    /// finite and non-negative. The previous delay is retained on error.
    pub fn set_autoplay_build_delay(&mut self, value: Option<f64>) -> Result<(), Error> {
        validate_delay(value)?;
        self.autoplay_build_delay = value;
        Ok(())
    }

    /// Return whether the idle timer is active when explicitly stored.
    #[must_use]
    pub const fn idle_timer_active(&self) -> Option<bool> {
        self.idle_timer_active
    }

    /// Set or clear the idle-timer active override.
    pub const fn set_idle_timer_active(&mut self, value: Option<bool>) {
        self.idle_timer_active = value;
    }

    /// Return the idle-timer delay in seconds.
    #[must_use]
    pub const fn idle_timer_delay(&self) -> Option<f64> {
        self.idle_timer_delay
    }

    /// Set or clear the idle-timer delay.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDelay`] when the value is present but not
    /// finite and non-negative. The previous delay is retained on error.
    pub fn set_idle_timer_delay(&mut self, value: Option<f64>) -> Result<(), Error> {
        validate_delay(value)?;
        self.idle_timer_delay = value;
        Ok(())
    }

    /// Return whether the show starts automatically when opened.
    #[must_use]
    pub const fn automatically_plays_upon_open(&self) -> Option<bool> {
        self.automatically_plays_upon_open
    }

    /// Set or clear the automatic-play-on-open override.
    pub const fn set_automatically_plays_upon_open(&mut self, value: Option<bool>) {
        self.automatically_plays_upon_open = value;
    }

    /// Validate every invariant, including values created through enum literals.
    ///
    /// # Errors
    ///
    /// Returns the first invalid-dimensions, non-canonical-mode, or invalid-
    /// delay error found.
    pub fn validate(&self) -> Result<(), Error> {
        let _ = Size::new(self.size.width, self.size.height)?;
        if self.mode.is_some_and(|mode| !mode.is_canonical()) {
            return Err(Error::NonCanonicalMode);
        }
        validate_delay(self.autoplay_transition_delay)?;
        validate_delay(self.autoplay_build_delay)?;
        validate_delay(self.idle_timer_delay)?;
        Ok(())
    }
}

impl Default for Settings {
    fn default() -> Self {
        Self {
            size: Size {
                width: 1024.0,
                height: 768.0,
            },
            slide_numbers_visible: None,
            loop_presentation: None,
            mode: None,
            autoplay_transition_delay: None,
            autoplay_build_delay: None,
            idle_timer_active: None,
            idle_timer_delay: None,
            automatically_plays_upon_open: None,
        }
    }
}

fn validate_delay(value: Option<f64>) -> Result<(), Error> {
    if value.is_some_and(|delay| !delay.is_finite() || delay < 0.0) {
        Err(Error::InvalidDelay)
    } else {
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, Mode, Settings, Size};

    #[test]
    fn size_rejects_non_positive_or_non_finite_dimensions() {
        assert_eq!(Size::new(0.0, 1.0), Err(Error::InvalidDimensions));
        assert_eq!(Size::new(f32::NAN, 1.0), Err(Error::InvalidDimensions));
        assert_eq!(Size::new(1.0, f32::INFINITY), Err(Error::InvalidDimensions));
    }

    #[test]
    fn mode_round_trips_unknown_values_without_aliasing_known_modes() {
        for (raw, mode) in [
            (0, Mode::Normal),
            (1, Mode::SelfPlaying),
            (2, Mode::LinksOnly),
            (19, Mode::Unknown(19)),
            (-1, Mode::Unknown(-1)),
        ] {
            assert_eq!(Mode::from_raw(raw), mode);
            assert_eq!(mode.as_raw(), raw);
        }
        assert_eq!(Mode::unknown(1), Err(Error::NonCanonicalMode));
        assert_eq!(Mode::unknown(19), Ok(Mode::Unknown(19)));
    }

    #[test]
    fn settings_mutators_preserve_invariants() {
        let mut settings = Settings::new(1920.0, 1080.0)
            .unwrap_or_else(|error| panic!("valid dimensions rejected: {error}"));
        settings
            .set_mode(Some(Mode::SelfPlaying))
            .unwrap_or_else(|error| panic!("valid mode rejected: {error}"));
        settings
            .set_autoplay_build_delay(Some(1.25))
            .unwrap_or_else(|error| panic!("valid delay rejected: {error}"));
        assert!((settings.size().width() - 1920.0).abs() < f32::EPSILON);
        assert_eq!(settings.mode(), Some(Mode::SelfPlaying));
        assert_eq!(
            settings.set_autoplay_build_delay(Some(f64::NAN)),
            Err(Error::InvalidDelay)
        );
        assert_eq!(
            settings.set_mode(Some(Mode::Unknown(1))),
            Err(Error::NonCanonicalMode)
        );
    }
}
