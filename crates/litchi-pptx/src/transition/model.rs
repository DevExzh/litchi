//! Semantic slide-transition values.

use std::fmt;
use std::sync::Arc;

/// Largest millisecond value accepted by `PowerPoint`'s transition timing
/// attributes.
///
/// Microsoft documents `advTm` as the inclusive range `0..=2_147_483_647`.
/// The same conservative bound is used for the Office 2010 transition
/// duration extension so both timing values remain accepted by `PowerPoint`.
pub const MAX_MS: u32 = i32::MAX as u32;

/// A checked `PowerPoint` transition time in milliseconds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[must_use]
pub struct Ms(u32);

impl Ms {
    /// Creates a checked millisecond value.
    ///
    /// # Errors
    ///
    /// Returns an error if `value` exceeds [`MAX_MS`].
    pub const fn new(value: u32) -> Result<Self, TimeError> {
        if value <= MAX_MS {
            Ok(Self(value))
        } else {
            Err(TimeError { value })
        }
    }

    /// Returns the encoded millisecond value.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }

    pub(crate) const fn known(value: u32) -> Self {
        Self(value)
    }
}

impl TryFrom<u32> for Ms {
    type Error = TimeError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Ms> for u32 {
    fn from(value: Ms) -> Self {
        value.get()
    }
}

/// A millisecond value lies outside `PowerPoint`'s checked domain.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TimeError {
    value: u32,
}

impl TimeError {
    /// Returns the rejected value.
    #[must_use]
    pub const fn value(self) -> u32 {
        self.value
    }
}

impl fmt::Display for TimeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "transition time {}ms exceeds the PowerPoint maximum of {MAX_MS}ms",
            self.value
        )
    }
}

impl std::error::Error for TimeError {}

/// A side of a slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Side {
    /// Left side.
    Left,
    /// Right side.
    Right,
    /// Top side.
    Up,
    /// Bottom side.
    Down,
}

/// A slide axis.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Axis {
    /// Horizontal axis.
    Horizontal,
    /// Vertical axis.
    Vertical,
}

/// A corner of a slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Corner {
    /// Upper-left corner.
    LeftUp,
    /// Upper-right corner.
    RightUp,
    /// Lower-left corner.
    LeftDown,
    /// Lower-right corner.
    RightDown,
}

/// An edge or corner used by cover and uncover effects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Origin {
    /// Left side.
    Left,
    /// Right side.
    Right,
    /// Top side.
    Up,
    /// Bottom side.
    Down,
    /// Upper-left corner.
    LeftUp,
    /// Upper-right corner.
    RightUp,
    /// Lower-left corner.
    LeftDown,
    /// Lower-right corner.
    RightDown,
}

/// An inward or outward movement.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum InOut {
    /// Move toward the center.
    In,
    /// Move away from the center.
    Out,
}

/// A geometry used by a shape transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Shape {
    /// Circle geometry.
    Circle,
    /// Diamond geometry.
    Diamond,
    /// Plus geometry.
    Plus,
}

/// Origin of a `PowerPoint` 2010 ripple effect.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Ripple {
    /// Center of the slide.
    Center,
    /// Upper-left corner.
    LeftUp,
    /// Upper-right corner.
    RightUp,
    /// Lower-left corner.
    LeftDown,
    /// Lower-right corner.
    RightDown,
}

/// PowerPoint-supported wheel spoke counts.
///
/// Unlike an integer field, this enum cannot represent spoke counts that
/// `PowerPoint` rejects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Spokes {
    /// One spoke.
    One = 1,
    /// Two spokes.
    Two = 2,
    /// Three spokes.
    Three = 3,
    /// Four spokes.
    Four = 4,
    /// Eight spokes.
    Eight = 8,
}

impl Spokes {
    /// Returns the encoded spoke count.
    #[must_use]
    pub const fn get(self) -> u8 {
        self as u8
    }
}

/// Transition speed preset.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Speed {
    /// Slow transition, nominally 1500ms.
    Slow,
    /// Medium transition, nominally 1000ms.
    #[default]
    Medium,
    /// Fast transition, nominally 500ms.
    Fast,
}

impl Speed {
    /// Returns the preset's nominal duration.
    pub const fn duration(self) -> Ms {
        match self {
            Self::Slow => Ms::known(1500),
            Self::Medium => Ms::known(1000),
            Self::Fast => Ms::known(500),
        }
    }
}

/// A validated, parsed-only transition child that has no semantic model yet.
///
/// `Raw` has no public constructor. Callers can inspect or clone a value read
/// from a document, but cannot inject unchecked XML through the safe facade.
/// A child that relies on a nonstandard namespace declaration outside its
/// captured subtree remains inspectable but is rejected by the writer.
#[derive(Clone)]
pub struct Raw {
    pub(crate) xml: Arc<str>,
    pub(crate) portable: bool,
}

impl Raw {
    /// Returns the retained XML subtree.
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Whether the subtree is self-contained or uses only namespace prefixes
    /// guaranteed by a generated slide root.
    #[must_use]
    pub const fn is_portable(&self) -> bool {
        self.portable
    }
}

impl fmt::Debug for Raw {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Raw")
            .field("bytes", &self.xml.len())
            .field("portable", &self.portable)
            .finish_non_exhaustive()
    }
}

impl PartialEq for Raw {
    fn eq(&self, other: &Self) -> bool {
        self.xml == other.xml && self.portable == other.portable
    }
}

impl Eq for Raw {}

/// A transition effect.
///
/// Each effect carries only the direction or option types accepted by its
/// `PresentationML` grammar. Invalid combinations therefore cannot be built.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    /// A transition element with no visual effect.
    None,
    /// Instant cut, optionally through black.
    Cut { black: Option<bool> },
    /// Fade, optionally through black.
    Fade { black: Option<bool> },
    /// Push from a slide side.
    Push(Side),
    /// Wipe from a slide side.
    Wipe(Side),
    /// Split along an axis, optionally moving in or out.
    Split { axis: Axis, toward: Option<InOut> },
    /// Pull the old slide away from an edge or corner.
    Uncover(Origin),
    /// Cover the old slide from an edge or corner.
    Cover(Origin),
    /// Dissolve pixels between slides.
    Dissolve,
    /// Alternating blinds along an axis.
    Blinds(Axis),
    /// Checkerboard along an axis.
    Checker(Axis),
    /// Random bars along an axis.
    RandomBars(Axis),
    /// Circle, diamond, or plus shape.
    Shape(Shape),
    /// Wedge sweep.
    Wedge,
    /// Zoom in or out.
    Zoom(InOut),
    /// Application-selected random effect.
    Random,
    /// Wheel with a PowerPoint-supported spoke count.
    Wheel(Spokes),
    /// Newsflash effect.
    Newsflash,
    /// `PowerPoint` 2010 ripple with a standard fade fallback.
    Ripple(Ripple),
    /// Diagonal strips from a corner.
    Strips(Corner),
    /// Comb along an axis.
    Comb(Axis),
    /// Validated, parsed-only effect retained for checked round-tripping when
    /// its namespace bindings are portable.
    Raw(Raw),
}

/// A complete slide-transition value.
#[derive(Debug, Clone)]
#[must_use]
pub struct Transition {
    pub(crate) kind: Kind,
    pub(crate) speed: Speed,
    pub(crate) duration: Option<Ms>,
    pub(crate) click: bool,
    pub(crate) after: Option<Ms>,
    pub(crate) preserved: Option<Arc<Preserved>>,
}

#[derive(Debug, Clone)]
pub(crate) struct Preserved {
    pub(crate) effect: Option<Raw>,
    pub(crate) before: Box<[Raw]>,
    pub(crate) after: Box<[Raw]>,
}

impl PartialEq for Transition {
    fn eq(&self, other: &Self) -> bool {
        self.same_semantics(other)
            && self.effect_xml() == other.effect_xml()
            && self.before() == other.before()
            && self.after_effect() == other.after_effect()
    }
}

impl Eq for Transition {}

impl Transition {
    /// Creates a transition with Office defaults.
    pub fn new(kind: Kind) -> Self {
        Self {
            kind,
            speed: Speed::Medium,
            duration: None,
            click: true,
            after: None,
            preserved: None,
        }
    }

    /// Returns the visual effect.
    #[must_use]
    pub fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Whether two values have the same modeled playback semantics.
    ///
    /// Unlike [`PartialEq`], this deliberately ignores retained raw XML. Use
    /// ordinary equality for no-op detection and exact authoring decisions.
    #[must_use]
    pub fn same_semantics(&self, other: &Self) -> bool {
        self.kind == other.kind
            && self.speed == other.speed
            && self.duration == other.duration
            && self.click == other.click
            && self.after == other.after
    }

    /// Replaces the visual effect.
    ///
    /// Parsed extension children remain attached, but raw XML for the old
    /// effect is discarded so the new typed value is serialized.
    pub fn set_kind(&mut self, kind: Kind) {
        self.kind = kind;
        self.preserved = self.preserved.take().and_then(|mut preserved| {
            let raw = Arc::make_mut(&mut preserved);
            raw.effect = None;
            (!raw.before.is_empty() || !raw.after.is_empty()).then_some(preserved)
        });
    }

    /// Replaces the visual effect in a builder chain.
    pub fn with_kind(mut self, kind: Kind) -> Self {
        self.set_kind(kind);
        self
    }

    /// Returns the speed preset.
    #[must_use]
    pub const fn speed(&self) -> Speed {
        self.speed
    }

    /// Sets the speed preset.
    pub fn set_speed(&mut self, speed: Speed) {
        self.speed = speed;
    }

    /// Sets the speed preset in a builder chain.
    pub fn with_speed(mut self, speed: Speed) -> Self {
        self.set_speed(speed);
        self
    }

    /// Returns the custom duration, if present.
    #[must_use]
    pub const fn duration(&self) -> Option<Ms> {
        self.duration
    }

    /// Sets or clears the custom duration.
    pub fn set_duration(&mut self, duration: Option<Ms>) {
        self.duration = duration;
    }

    /// Sets a custom duration in a builder chain.
    pub fn with_duration(mut self, duration: Ms) -> Self {
        self.set_duration(Some(duration));
        self
    }

    /// Returns whether a click advances the slide.
    #[must_use]
    pub const fn click(&self) -> bool {
        self.click
    }

    /// Enables or disables click-to-advance.
    pub fn set_click(&mut self, click: bool) {
        self.click = click;
    }

    /// Sets click-to-advance in a builder chain.
    pub fn with_click(mut self, click: bool) -> Self {
        self.set_click(click);
        self
    }

    /// Returns the automatic-advance delay, if present.
    #[must_use]
    pub const fn after(&self) -> Option<Ms> {
        self.after
    }

    /// Sets or clears automatic advance.
    pub fn set_after(&mut self, after: Option<Ms>) {
        self.after = after;
    }

    /// Sets automatic advance in a builder chain.
    pub fn with_after(mut self, after: Ms) -> Self {
        self.set_after(Some(after));
        self
    }

    /// Returns the effective duration, preferring a custom duration.
    pub const fn effective_duration(&self) -> Ms {
        match self.duration {
            Some(duration) => duration,
            None => self.speed.duration(),
        }
    }

    /// Iterates over inert extension children retained around the effect.
    pub fn preserved(&self) -> impl Iterator<Item = &Raw> {
        self.before().iter().chain(self.after_effect().iter())
    }

    /// Returns the number of inert extension children retained around the
    /// effect.
    #[must_use]
    pub fn preserved_len(&self) -> usize {
        self.before()
            .len()
            .saturating_add(self.after_effect().len())
    }

    pub(crate) fn effect_xml(&self) -> Option<&Raw> {
        self.preserved
            .as_deref()
            .and_then(|preserved| preserved.effect.as_ref())
    }

    pub(crate) fn before(&self) -> &[Raw] {
        self.preserved
            .as_deref()
            .map_or(&[], |preserved| preserved.before.as_ref())
    }

    pub(crate) fn after_effect(&self) -> &[Raw] {
        self.preserved
            .as_deref()
            .map_or(&[], |preserved| preserved.after.as_ref())
    }
}

pub(crate) fn preserved_effect_xml(value: &Transition) -> Option<&str> {
    value.effect_xml().map(Raw::xml)
}

pub(crate) fn semantic_clone(value: &Transition) -> Transition {
    let mut value = value.clone();
    value.preserved = None;
    value
}
