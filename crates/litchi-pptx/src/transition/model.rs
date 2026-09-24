//! Semantic slide-transition values.

use std::fmt;
use std::sync::Arc;

use crate::time::Offset;
use crate::{Error, Result as PptxResult};

/// Largest millisecond value accepted by the legacy integral transition
/// timing attributes.
///
/// Microsoft documents `advTm` as the inclusive range `0..=2_147_483_647`.
/// Exact Office 2010 `p14:dur` values use [`crate::time::Offset`] and are
/// bounded by that type's lexical safety limit instead.
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

/// A horizontal side used by the PowerPoint 2010 effects whose schema only
/// accepts `l` or `r`.
///
/// This is deliberately separate from [`Side`].  Passing `Up` or `Down` to a
/// conveyor, ferris, flip, gallery, reveal, or switch effect would produce an
/// invalid `ST_TransitionLeftRightDirectionType` value, so the narrower type
/// keeps those combinations out of the authoring API.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LeftRight {
    /// Left.
    Left,
    /// Right.
    Right,
    /// The optional `dir` attribute was omitted.
    ///
    /// This is accepted by the `CT_LeftRightDirectionTransition` effects.
    /// `Reveal` also permits the omitted form and applies its schema default
    /// when played by PowerPoint.
    Unspecified,
}

impl LeftRight {
    pub(crate) const fn wire(self) -> Option<&'static str> {
        match self {
            Self::Left => Some("l"),
            Self::Right => Some("r"),
            Self::Unspecified => None,
        }
    }
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

/// A geometric pattern used by a PowerPoint 2010 glitter transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum GlitterPattern {
    /// Diamond tiles.
    Diamond,
    /// Hexagonal tiles.
    Hexagon,
}

impl GlitterPattern {
    pub(crate) const fn wire(self) -> &'static str {
        match self {
            Self::Diamond => "diamond",
            Self::Hexagon => "hexagon",
        }
    }
}

/// A geometric pattern used by a PowerPoint 2010 shred transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ShredPattern {
    /// Vertical strips.
    Strip,
    /// Small rectangles.
    Rectangle,
}

impl ShredPattern {
    pub(crate) const fn wire(self) -> &'static str {
        match self {
            Self::Strip => "strip",
            Self::Rectangle => "rectangle",
        }
    }
}

/// Parameters for a PowerPoint 2010 fly-through transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct FlyThrough {
    direction: InOut,
    bounce: bool,
}

impl FlyThrough {
    /// Creates a fly-through with its direction and bounce setting.
    #[must_use]
    pub const fn new(direction: InOut, bounce: bool) -> Self {
        Self { direction, bounce }
    }

    /// Returns the slide movement direction.
    #[must_use]
    pub const fn direction(self) -> InOut {
        self.direction
    }

    /// Returns whether the slide movement bounces.
    #[must_use]
    pub const fn bounce(self) -> bool {
        self.bounce
    }

    /// Sets the slide movement direction.
    pub const fn set_direction(&mut self, direction: InOut) {
        self.direction = direction;
    }

    /// Sets the slide movement direction in a builder chain.
    #[must_use]
    pub const fn with_direction(mut self, direction: InOut) -> Self {
        self.set_direction(direction);
        self
    }

    /// Sets whether the slide movement bounces.
    pub const fn set_bounce(&mut self, bounce: bool) {
        self.bounce = bounce;
    }

    /// Sets the bounce flag in a builder chain.
    #[must_use]
    pub const fn with_bounce(mut self, bounce: bool) -> Self {
        self.set_bounce(bounce);
        self
    }
}

/// Parameters for a PowerPoint 2010 glitter transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Glitter {
    direction: Side,
    pattern: GlitterPattern,
}

impl Glitter {
    /// Creates a glitter transition with a direction and tile pattern.
    #[must_use]
    pub const fn new(direction: Side, pattern: GlitterPattern) -> Self {
        Self { direction, pattern }
    }

    /// Returns the slide movement direction.
    #[must_use]
    pub const fn direction(self) -> Side {
        self.direction
    }

    /// Returns the tile pattern.
    #[must_use]
    pub const fn pattern(self) -> GlitterPattern {
        self.pattern
    }

    /// Sets the slide movement direction.
    pub const fn set_direction(&mut self, direction: Side) {
        self.direction = direction;
    }

    /// Sets the slide movement direction in a builder chain.
    #[must_use]
    pub const fn with_direction(mut self, direction: Side) -> Self {
        self.set_direction(direction);
        self
    }

    /// Sets the tile pattern.
    pub const fn set_pattern(&mut self, pattern: GlitterPattern) {
        self.pattern = pattern;
    }

    /// Sets the tile pattern in a builder chain.
    #[must_use]
    pub const fn with_pattern(mut self, pattern: GlitterPattern) -> Self {
        self.set_pattern(pattern);
        self
    }
}

/// Parameters for a PowerPoint 2010 prism transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Prism {
    direction: Side,
    content: bool,
    inverted: bool,
}

impl Prism {
    /// Creates a prism transition with its direction and rendering flags.
    #[must_use]
    pub const fn new(direction: Side, content: bool, inverted: bool) -> Self {
        Self {
            direction,
            content,
            inverted,
        }
    }

    /// Returns the slide movement direction.
    #[must_use]
    pub const fn direction(self) -> Side {
        self.direction
    }

    /// Returns whether content and the background are drawn separately.
    #[must_use]
    pub const fn content(self) -> bool {
        self.content
    }

    /// Returns whether the prism is concave.
    #[must_use]
    pub const fn inverted(self) -> bool {
        self.inverted
    }

    /// Sets the slide movement direction.
    pub const fn set_direction(&mut self, direction: Side) {
        self.direction = direction;
    }

    /// Sets the slide movement direction in a builder chain.
    #[must_use]
    pub const fn with_direction(mut self, direction: Side) -> Self {
        self.set_direction(direction);
        self
    }

    /// Sets whether content and the background are drawn separately.
    pub const fn set_content(&mut self, content: bool) {
        self.content = content;
    }

    /// Sets the content-separation flag in a builder chain.
    #[must_use]
    pub const fn with_content(mut self, content: bool) -> Self {
        self.set_content(content);
        self
    }

    /// Sets whether the prism is concave.
    pub const fn set_inverted(&mut self, inverted: bool) {
        self.inverted = inverted;
    }

    /// Sets the concavity flag in a builder chain.
    #[must_use]
    pub const fn with_inverted(mut self, inverted: bool) -> Self {
        self.set_inverted(inverted);
        self
    }
}

/// Parameters for a PowerPoint 2010 reveal transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Reveal {
    direction: LeftRight,
    through_black: bool,
}

impl Reveal {
    /// Creates a reveal transition with its direction and black-screen flag.
    #[must_use]
    pub const fn new(direction: LeftRight, through_black: bool) -> Self {
        Self {
            direction: normalize_reveal_direction(direction),
            through_black,
        }
    }

    /// Returns the slide movement direction.
    #[must_use]
    pub const fn direction(self) -> LeftRight {
        self.direction
    }

    /// Returns whether the transition fades through black.
    #[must_use]
    pub const fn through_black(self) -> bool {
        self.through_black
    }

    /// Sets the slide movement direction.
    pub const fn set_direction(&mut self, direction: LeftRight) {
        self.direction = normalize_reveal_direction(direction);
    }

    /// Sets the slide movement direction in a builder chain.
    #[must_use]
    pub const fn with_direction(mut self, direction: LeftRight) -> Self {
        self.set_direction(direction);
        self
    }

    /// Sets whether the transition fades through black.
    pub const fn set_through_black(&mut self, through_black: bool) {
        self.through_black = through_black;
    }

    /// Sets the black-screen flag in a builder chain.
    #[must_use]
    pub const fn with_through_black(mut self, through_black: bool) -> Self {
        self.set_through_black(through_black);
        self
    }
}

/// Parameters for a PowerPoint 2010 shred transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Shred {
    pattern: ShredPattern,
    direction: InOut,
}

impl Shred {
    /// Creates a shred transition with its tile pattern and direction.
    #[must_use]
    pub const fn new(pattern: ShredPattern, direction: InOut) -> Self {
        Self { pattern, direction }
    }

    /// Returns the tile pattern.
    #[must_use]
    pub const fn pattern(self) -> ShredPattern {
        self.pattern
    }

    /// Returns the slide movement direction.
    #[must_use]
    pub const fn direction(self) -> InOut {
        self.direction
    }

    /// Sets the tile pattern.
    pub const fn set_pattern(&mut self, pattern: ShredPattern) {
        self.pattern = pattern;
    }

    /// Sets the tile pattern in a builder chain.
    #[must_use]
    pub const fn with_pattern(mut self, pattern: ShredPattern) -> Self {
        self.set_pattern(pattern);
        self
    }

    /// Sets the slide movement direction.
    pub const fn set_direction(&mut self, direction: InOut) {
        self.direction = direction;
    }

    /// Sets the slide movement direction in a builder chain.
    #[must_use]
    pub const fn with_direction(mut self, direction: InOut) -> Self {
        self.set_direction(direction);
        self
    }
}

/// Matching granularity for a PowerPoint Morph transition.
///
/// The values are the complete `ST_TransitionMorphOption` vocabulary from
/// `[MS-PPTX]` 2.6.4.1. Keeping this as a closed value prevents a malformed
/// token from being emitted by the ordinary transition writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Morph {
    /// Match and move whole objects.
    ByObject,
    /// Match and move objects and individual words.
    ByWord,
    /// Match and move objects and individual characters.
    ByChar,
}

/// Maximum UTF-8 bytes retained for one preset transition name.
pub const MAX_PRESET_NAME_BYTES: usize = 64 * 1024;

impl Morph {
    pub(crate) const fn wire(self) -> &'static str {
        match self {
            Self::ByObject => "byObject",
            Self::ByWord => "byWord",
            Self::ByChar => "byChar",
        }
    }
}

/// A typed PowerPoint preset transition (`p15:prstTrans`).
///
/// `prst` is intentionally retained as a checked string because the
/// Microsoft schema declares it as `xsd:string`, while the specification only
/// documents the presets known to a particular Office release. This keeps
/// newer producer names readable and writable without guessing their visual
/// meaning. The optional value also preserves the schema's omitted-attribute
/// form.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Preset {
    name: Option<Box<str>>,
    invert_x: bool,
    invert_y: bool,
}

impl Preset {
    /// Creates a preset with the supplied `prst` name and default inversion.
    ///
    /// The value is checked for XML 1.0 characters before it is retained.
    ///
    /// # Errors
    ///
    /// Returns an error when the name is too large or contains an invalid XML
    /// 1.0 character.
    pub fn new(name: impl AsRef<str>) -> PptxResult<Self> {
        Self::with_options(name.as_ref(), false, false)
    }

    /// Creates a preset with all three schema attributes.
    ///
    /// # Errors
    ///
    /// Returns an error when the name is too large or contains an invalid XML
    /// 1.0 character.
    pub fn with_options(name: impl AsRef<str>, invert_x: bool, invert_y: bool) -> PptxResult<Self> {
        Ok(Self {
            name: Some(validate_preset_name(name.as_ref())?),
            invert_x,
            invert_y,
        })
    }

    /// Creates the schema-valid form with an omitted `prst` attribute.
    #[must_use]
    pub const fn without_name() -> Self {
        Self {
            name: None,
            invert_x: false,
            invert_y: false,
        }
    }

    /// Returns the optional `prst` name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the `invX` value.
    #[must_use]
    pub const fn invert_x(&self) -> bool {
        self.invert_x
    }

    /// Returns the `invY` value.
    #[must_use]
    pub const fn invert_y(&self) -> bool {
        self.invert_y
    }

    /// Sets whether the preset's X coordinates are inverted.
    pub fn set_invert_x(&mut self, value: bool) {
        self.invert_x = value;
    }

    /// Sets whether the preset's Y coordinates are inverted.
    pub fn set_invert_y(&mut self, value: bool) {
        self.invert_y = value;
    }

    /// Sets X inversion in a builder chain.
    pub fn with_invert_x(mut self, value: bool) -> Self {
        self.set_invert_x(value);
        self
    }

    /// Sets Y inversion in a builder chain.
    pub fn with_invert_y(mut self, value: bool) -> Self {
        self.set_invert_y(value);
        self
    }

    pub(crate) fn from_parts(
        name: Option<String>,
        invert_x: bool,
        invert_y: bool,
    ) -> PptxResult<Self> {
        Ok(Self {
            name: name.map(validate_owned_preset_name).transpose()?,
            invert_x,
            invert_y,
        })
    }
}

impl Default for Preset {
    fn default() -> Self {
        Self::without_name()
    }
}

fn validate_preset_name(name: &str) -> PptxResult<Box<str>> {
    validate_preset_name_bytes(name)?;
    Ok(name.into())
}

fn validate_owned_preset_name(name: String) -> PptxResult<Box<str>> {
    validate_preset_name_bytes(&name)?;
    Ok(name.into_boxed_str())
}

fn validate_preset_name_bytes(name: &str) -> PptxResult<()> {
    if name.len() > MAX_PRESET_NAME_BYTES {
        return Err(Error::Limit {
            resource: "preset transition name bytes",
            limit: MAX_PRESET_NAME_BYTES,
        });
    }
    if name.chars().any(|character| {
        let value = character as u32;
        !matches!(
            value,
            0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
    }) {
        return Err(Error::Invalid(
            "preset transition name contains an invalid XML character".into(),
        ));
    }
    Ok(())
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
    /// PowerPoint 2010 conveyor transition.
    Conveyor(LeftRight),
    /// PowerPoint 2010 doors transition.
    Doors(Axis),
    /// PowerPoint 2010 ferris transition.
    Ferris(LeftRight),
    /// PowerPoint 2010 flash transition.
    Flash,
    /// PowerPoint 2010 flip transition.
    Flip(LeftRight),
    /// PowerPoint 2010 fly-through transition.
    FlyThrough(FlyThrough),
    /// PowerPoint 2010 gallery transition.
    Gallery(LeftRight),
    /// PowerPoint 2010 glitter transition.
    Glitter(Glitter),
    /// PowerPoint 2010 honeycomb transition.
    Honeycomb,
    /// PowerPoint 2010 pan transition.
    Pan(Side),
    /// PowerPoint 2010 prism transition.
    Prism(Prism),
    /// PowerPoint 2010 reveal transition.
    Reveal(Reveal),
    /// PowerPoint 2010 shred transition.
    Shred(Shred),
    /// PowerPoint 2010 switch transition.
    Switch(LeftRight),
    /// PowerPoint 2010 vortex transition.
    Vortex(Side),
    /// PowerPoint 2010 warp transition.
    Warp(InOut),
    /// PowerPoint 2010 reverse-wheel transition.
    WheelReverse(Spokes),
    /// PowerPoint 2010 window transition.
    Window(Axis),
    /// PowerPoint Morph transition with a standard fade fallback.
    Morph(Morph),
    /// PowerPoint 2012 preset transition with a standard fade fallback.
    Preset(Preset),
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
    pub(crate) duration_offset: Option<Arc<Offset>>,
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
            duration_offset: None,
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
            && self.duration_offset == other.duration_offset
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

    /// Returns an integral custom duration when it fits the v1 millisecond
    /// domain. Fractional or larger exact values are available through
    /// [`Self::duration_offset`].
    #[must_use]
    pub fn duration(&self) -> Option<Ms> {
        self.duration_offset.as_deref().and_then(offset_as_ms)
    }

    /// Sets or clears the custom duration.
    pub fn set_duration(&mut self, duration: Option<Ms>) {
        self.set_duration_offset(duration.map(|value| Offset::ms(u64::from(value.get()))));
    }

    /// Sets a custom duration in a builder chain.
    pub fn with_duration(mut self, duration: Ms) -> Self {
        self.set_duration(Some(duration));
        self
    }

    /// Returns the exact custom duration, if present.
    ///
    /// The value is normalized to decimal milliseconds, but fractional
    /// milliseconds are retained. Use this accessor when handling the
    /// Office 2010 `p14:dur` extension; [`Self::duration`] remains the v1
    /// integral-millisecond view.
    #[must_use]
    pub fn duration_offset(&self) -> Option<&Offset> {
        self.duration_offset.as_deref()
    }

    /// Sets or clears the exact custom duration.
    ///
    /// Integral values that fit the v1 [`Ms`] domain are also exposed through
    /// [`Self::duration`]. Larger or fractional values remain available only
    /// through [`Self::duration_offset`] so no precision is discarded.
    pub fn set_duration_offset(&mut self, duration: Option<Offset>) {
        self.duration_offset = duration.map(Arc::new);
    }

    /// Sets an exact custom duration in a builder chain.
    pub fn with_duration_offset(mut self, duration: Offset) -> Self {
        self.set_duration_offset(Some(duration));
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

    /// Returns the effective integral duration, preferring a custom duration
    /// that fits the v1 millisecond domain. Use
    /// [`Self::effective_duration_offset`] for exact p14 timing.
    pub fn effective_duration(&self) -> Ms {
        match self.duration() {
            Some(duration) => duration,
            None => self.speed.duration(),
        }
    }

    /// Returns the exact effective duration, including fractional
    /// milliseconds from `p14:dur`.
    pub fn effective_duration_offset(&self) -> Offset {
        self.duration_offset
            .as_deref()
            .cloned()
            .unwrap_or_else(|| Offset::ms(u64::from(self.speed.duration().get())))
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

fn offset_as_ms(offset: &Offset) -> Option<Ms> {
    let value = offset.as_str();
    if value.contains('.') {
        return None;
    }
    let value = value.parse::<u64>().ok()?;
    let value = u32::try_from(value).ok()?;
    Ms::new(value).ok()
}

const fn normalize_reveal_direction(direction: LeftRight) -> LeftRight {
    match direction {
        LeftRight::Unspecified => LeftRight::Left,
        direction => direction,
    }
}
