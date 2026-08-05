//! Keynote show settings and immutable presentation snapshots.

use crate::{Error, Result, Seconds, Slide};

const NORMAL_MODE: i32 = 0;
const SELF_PLAYING_MODE: i32 = 1;
const LINKS_ONLY_MODE: i32 = 2;

/// A finite, strictly positive presentation size in points.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Size {
    width: f32,
    height: f32,
}

impl Size {
    /// Construct a validated presentation size.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDimensions`] when either dimension is not
    /// finite and strictly positive.
    pub fn new(width: f32, height: f32) -> Result<Self> {
        if width.is_finite() && width > 0.0 && height.is_finite() && height > 0.0 {
            Ok(Self { width, height })
        } else {
            Err(Error::InvalidDimensions)
        }
    }

    /// Return the width in points.
    #[must_use]
    pub const fn width(self) -> f32 {
        self.width
    }

    /// Return the height in points.
    #[must_use]
    pub const fn height(self) -> f32 {
        self.height
    }
}

impl Default for Size {
    fn default() -> Self {
        Self {
            width: 1_024.0,
            height: 768.0,
        }
    }
}

/// How a presentation advances through its slides.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Mode {
    /// Advance through presenter input and configured transitions.
    Normal,
    /// Advance automatically using playback delays.
    SelfPlaying,
    /// Navigate only through hyperlinks.
    LinksOnly,
    /// A mode introduced by a newer Keynote release.
    Unknown(i32),
}

impl Mode {
    /// Decode a native mode value without discarding unknown values.
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        match raw {
            NORMAL_MODE => Self::Normal,
            SELF_PLAYING_MODE => Self::SelfPlaying,
            LINKS_ONLY_MODE => Self::LinksOnly,
            other => Self::Unknown(other),
        }
    }

    /// Construct an unknown mode while rejecting values assigned to a named
    /// native variant.
    pub const fn unknown(raw: i32) -> Result<Self> {
        match raw {
            NORMAL_MODE | SELF_PLAYING_MODE | LINKS_ONLY_MODE => {
                Err(Error::NonCanonicalMode)
            }
            other => Ok(Self::Unknown(other)),
        }
    }

    /// Return the native mode value.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Normal => NORMAL_MODE,
            Self::SelfPlaying => SELF_PLAYING_MODE,
            Self::LinksOnly => LINKS_ONLY_MODE,
            Self::Unknown(value) => value,
        }
    }

    /// Return whether this value uses a named variant for known input.
    #[must_use]
    pub const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(NORMAL_MODE | SELF_PLAYING_MODE | LINKS_ONLY_MODE)
        )
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
struct Flags {
    present: u8,
    values: u8,
}

#[derive(Debug, Clone, Copy)]
enum Flag {
    SlideNumbersVisible,
    LoopPresentation,
    IdleTimerActive,
    AutomaticallyPlaysUponOpen,
}

impl Flag {
    const fn bit(self) -> u8 {
        match self {
            Self::SlideNumbersVisible => 1,
            Self::LoopPresentation => 2,
            Self::IdleTimerActive => 4,
            Self::AutomaticallyPlaysUponOpen => 8,
        }
    }
}

impl Flags {
    fn get(self, flag: Flag) -> Option<bool> {
        let bit = flag.bit();
        (self.present & bit != 0).then_some(self.values & bit != 0)
    }

    fn set(&mut self, flag: Flag, value: Option<bool>) {
        let bit = flag.bit();
        if let Some(set_value) = value {
            self.present |= bit;
            if set_value {
                self.values |= bit;
            } else {
                self.values &= !bit;
            }
        } else {
            self.present &= !bit;
            self.values &= !bit;
        }
    }
}

/// Validated show-level playback settings.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Settings {
    size: Size,
    flags: Flags,
    mode: Option<Mode>,
    autoplay_transition_delay: Option<Seconds>,
    autoplay_build_delay: Option<Seconds>,
    idle_timer_delay: Option<Seconds>,
}

impl Settings {
    /// Create settings with all optional producer fields absent.
    #[must_use]
    pub const fn new(size: Size) -> Self {
        Self {
            size,
            flags: Flags {
                present: 0,
                values: 0,
            },
            mode: None,
            autoplay_transition_delay: None,
            autoplay_build_delay: None,
            idle_timer_delay: None,
        }
    }

    /// Return the presentation size.
    #[must_use]
    pub const fn size(self) -> Size {
        self.size
    }

    /// Replace the presentation size.
    pub fn set_size(&mut self, size: Size) {
        self.size = size;
    }

    /// Return the optional slide-number visibility flag.
    #[must_use]
    pub fn slide_numbers_visible(self) -> Option<bool> {
        self.flags.get(Flag::SlideNumbersVisible)
    }

    /// Set or clear slide-number visibility.
    pub fn set_slide_numbers_visible(&mut self, value: Option<bool>) {
        self.flags.set(Flag::SlideNumbersVisible, value);
    }

    /// Return the optional looping flag.
    #[must_use]
    pub fn loop_presentation(self) -> Option<bool> {
        self.flags.get(Flag::LoopPresentation)
    }

    /// Set or clear looping.
    pub fn set_loop_presentation(&mut self, value: Option<bool>) {
        self.flags.set(Flag::LoopPresentation, value);
    }

    /// Return the optional playback mode.
    #[must_use]
    pub const fn mode(self) -> Option<Mode> {
        self.mode
    }

    /// Set or clear playback mode, rejecting non-canonical known values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalMode`] when a known native value is passed
    /// through [`Mode::Unknown`].
    pub fn set_mode(&mut self, mode: Option<Mode>) -> Result<()> {
        if mode.is_some_and(|value| !value.is_canonical()) {
            return Err(Error::NonCanonicalMode);
        }
        self.mode = mode;
        Ok(())
    }

    /// Return the automatic transition delay.
    #[must_use]
    pub const fn autoplay_transition_delay(self) -> Option<Seconds> {
        self.autoplay_transition_delay
    }

    /// Set or clear the automatic transition delay.
    pub fn set_autoplay_transition_delay(&mut self, value: Option<Seconds>) {
        self.autoplay_transition_delay = value;
    }

    /// Return the automatic build delay.
    #[must_use]
    pub const fn autoplay_build_delay(self) -> Option<Seconds> {
        self.autoplay_build_delay
    }

    /// Set or clear the automatic build delay.
    pub fn set_autoplay_build_delay(&mut self, value: Option<Seconds>) {
        self.autoplay_build_delay = value;
    }

    /// Return whether Keynote activates its idle timer.
    #[must_use]
    pub fn idle_timer_active(self) -> Option<bool> {
        self.flags.get(Flag::IdleTimerActive)
    }

    /// Set or clear idle-timer activation.
    pub fn set_idle_timer_active(&mut self, value: Option<bool>) {
        self.flags.set(Flag::IdleTimerActive, value);
    }

    /// Return the idle-timer delay.
    #[must_use]
    pub const fn idle_timer_delay(self) -> Option<Seconds> {
        self.idle_timer_delay
    }

    /// Set or clear the idle-timer delay.
    pub fn set_idle_timer_delay(&mut self, value: Option<Seconds>) {
        self.idle_timer_delay = value;
    }

    /// Return whether the presentation plays when opened.
    #[must_use]
    pub fn automatically_plays_upon_open(self) -> Option<bool> {
        self.flags.get(Flag::AutomaticallyPlaysUponOpen)
    }

    /// Set or clear automatic playback on open.
    pub fn set_automatically_plays_upon_open(&mut self, value: Option<bool>) {
        self.flags.set(Flag::AutomaticallyPlaysUponOpen, value);
    }

    /// Validate all semantic invariants.
    ///
    /// # Errors
    ///
    /// Returns a typed semantic error when the size or mode is invalid.
    pub fn validate(self) -> Result<()> {
        Size::new(self.size.width, self.size.height)?;
        if self.mode.is_some_and(|value| !value.is_canonical()) {
            return Err(Error::NonCanonicalMode);
        }
        Ok(())
    }
}

impl Default for Settings {
    fn default() -> Self {
        Self::new(Size::default())
    }
}

/// An immutable presentation snapshot.
#[derive(Debug, Clone, PartialEq)]
pub struct Show {
    title: Option<Box<str>>,
    slides: Box<[Slide]>,
    settings: Settings,
}

impl Show {
    /// Start a detached show builder.
    #[must_use]
    pub fn builder() -> Builder {
        Builder::new()
    }

    /// Return the optional presentation title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Borrow slides in presentation order.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Return the number of slides.
    #[must_use]
    pub const fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Select a slide by checked zero-based position.
    #[must_use]
    pub fn slide(&self, index: usize) -> Option<&Slide> {
        self.slides.get(index)
    }

    /// Borrow validated show settings.
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }

    /// Return all show and slide text in presentation order.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let mut text = Vec::new();
        if let Some(title) = &self.title {
            text.push(title.to_string());
        }
        for slide in &self.slides {
            text.extend(slide.all_text());
        }
        text
    }

    /// Return whether the show contains no slides.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.slides.is_empty()
    }
}

/// A detached, mutable show builder.
#[derive(Debug, Default)]
pub struct Builder {
    title: Option<Box<str>>,
    slides: Vec<Slide>,
    settings: Settings,
}

impl Builder {
    /// Create an empty show builder with standard settings.
    #[must_use]
    pub fn new() -> Self {
        Self {
            title: None,
            slides: Vec::new(),
            settings: Settings::default(),
        }
    }

    /// Set or clear the presentation title.
    pub fn set_title(&mut self, title: Option<String>) {
        self.title = title.map(String::into_boxed_str);
    }

    /// Append one slide in source order.
    pub fn push_slide(&mut self, slide: Slide) {
        self.slides.push(slide);
    }

    /// Replace the validated settings.
    pub fn set_settings(&mut self, settings: Settings) {
        self.settings = settings;
    }

    /// Finish the builder as an immutable show snapshot.
    #[must_use]
    pub fn build(self) -> Show {
        Show {
            title: self.title,
            slides: self.slides.into_boxed_slice(),
            settings: self.settings,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Slide;

    #[test]
    fn size_and_modes_are_checked_and_lossless() -> Result<()> {
        assert_eq!(Size::new(0.0, 1.0), Err(Error::InvalidDimensions));
        for raw in [0, 1, 2, 19, -1] {
            assert_eq!(Mode::from_raw(raw).as_raw(), raw);
        }
        let mut settings = Settings::default();
        settings.set_mode(Some(Mode::SelfPlaying))?;
        assert_eq!(settings.mode(), Some(Mode::SelfPlaying));
        assert_eq!(
            settings.set_mode(Some(Mode::Unknown(1))),
            Err(Error::NonCanonicalMode)
        );
        Ok(())
    }

    #[test]
    fn show_is_an_immutable_ordered_snapshot() {
        let mut builder = Show::builder();
        builder.set_title(Some("Deck".to_owned()));
        builder.push_slide(Slide::builder(0).build());
        let show = builder.build();
        assert_eq!(show.title(), Some("Deck"));
        assert_eq!(show.slide_count(), 1);
        assert_eq!(show.slide(0).map(Slide::index), Some(0));
        assert!(show.slide(1).is_none());
    }
}
