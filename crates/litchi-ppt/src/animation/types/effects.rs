//! Authoring-level effects and compatibility timeline types.

use crate::animation::motion_path::MotionPath;
use crate::animation::sound::AnimationSound;
use crate::animation::triggers::IterationType;
use crate::records::Record;

/// A single build level (animation step).
#[derive(Debug, Clone, PartialEq)]
pub struct BuildLevel {
    /// Build type (entrance, emphasis, exit, etc.)
    pub build_type: BuildType,
    /// Shape ID that is animated
    pub shape_id: u32,
    /// Build order (0-indexed)
    pub build_order: u32,
    /// Animation effect
    pub effect: AnimationEffect,
    /// Effect speed
    pub speed: EffectSpeed,
    /// Effect direction
    pub direction: EffectDirection,
    /// Trigger type
    pub trigger: AnimationTrigger,
    /// Motion path (if this is a motion path animation)
    pub motion_path: Option<MotionPath>,
    /// Sound for this animation
    pub sound: Option<AnimationSound>,
    /// Iteration type (for text)
    pub iteration: IterationType,
    /// After-effect behavior
    pub after_effect: AfterEffect,
    /// Duration override in milliseconds (None = use default for speed)
    pub duration_ms: Option<u32>,
}

impl Default for BuildLevel {
    fn default() -> Self {
        Self {
            build_type: BuildType::Entrance,
            shape_id: 0,
            build_order: 0,
            effect: AnimationEffect::Appear,
            speed: EffectSpeed::Medium,
            direction: EffectDirection::None,
            trigger: AnimationTrigger::OnClick,
            motion_path: None,
            sound: None,
            iteration: IterationType::default(),
            after_effect: AfterEffect::None,
            duration_ms: None,
        }
    }
}

/// Build type (animation category).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum BuildType {
    /// Entrance effect
    Entrance,
    /// Emphasis effect
    Emphasis,
    /// Exit effect
    Exit,
    /// Motion path
    MotionPath,
}

/// Animation effect type.
/// Covers entrance, emphasis, exit, and motion path effects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum AnimationEffect {
    // Entrance Effects
    /// Appear
    #[default]
    Appear,
    /// Fade in
    FadeIn,
    /// Fly in
    FlyIn,
    /// Wipe
    Wipe,
    /// Split
    Split,
    /// Dissolve
    Dissolve,
    /// Box
    Box,
    /// Checkerboard
    Checkerboard,
    /// Blinds
    Blinds,
    /// Random bars
    RandomBars,
    /// Grow and turn
    GrowAndTurn,
    /// Zoom
    Zoom,
    /// Swivel
    Swivel,
    /// Bounce
    Bounce,
    /// Float in
    FloatIn,
    /// Ascend
    Ascend,
    /// Descend
    Descend,
    /// Expand
    Expand,
    /// Compress
    Compress,
    /// Stretch
    Stretch,
    /// Wheel
    Wheel,
    /// Peek in
    PeekIn,
    /// Plus
    Plus,
    /// Diamond
    Diamond,
    /// Wedge
    Wedge,
    /// Strips
    Strips,
    /// Random
    Random,
    /// Crawl in
    CrawlIn,
    /// Rise up
    RiseUp,
    /// Spiral in
    SpiralIn,

    // Emphasis Effects
    /// Pulse
    Pulse,
    /// Spin
    Spin,
    /// Teeter
    Teeter,
    /// Wave
    Wave,
    /// Lighten
    Lighten,
    /// Darken
    Darken,
    /// Change fill color
    ChangeFillColor,
    /// Change line color
    ChangeLineColor,
    /// Change font color
    ChangeFontColor,
    /// Change font size
    ChangeFontSize,
    /// Grow/Shrink
    GrowShrink,
    /// Bold flash
    BoldFlash,
    /// Underline
    Underline,
    /// Color pulse
    ColorPulse,
    /// Complementary color
    ComplementaryColor,
    /// Complementary color 2
    ComplementaryColor2,
    /// Contrasting color
    ContrastingColor,
    /// Transparency
    Transparency,
    /// Object color
    ObjectColor,
    /// Vertical highlight
    VerticalHighlight,
    /// Flicker
    Flicker,

    // Exit Effects
    /// Fade out (same as `FadeIn` but exit type)
    FadeOut,
    /// Fly out
    FlyOut,
    /// Wipe out
    WipeOut,
    /// Disappear
    Disappear,
    /// Box out
    BoxOut,
    /// Checkerboard out
    CheckerboardOut,
    /// Blinds out
    BlindsOut,
    /// Random bars out
    RandomBarsOut,
    /// Strips out
    StripsOut,
    /// Split out
    SplitOut,
    /// Peek out
    PeekOut,
    /// Plus out
    PlusOut,
    /// Diamond out
    DiamondOut,
    /// Crawl out
    CrawlOut,
    /// Descend out
    DescendOut,
    /// Collapse
    Collapse,
    /// Sink down
    SinkDown,
    /// Spiral out
    SpiralOut,

    // Motion Path Effects
    /// Custom motion path
    MotionPath,
    /// Lines motion path
    MotionPathLines,
    /// Curves motion path
    MotionPathCurves,
    /// Shapes motion path
    MotionPathShapes,
    /// Left motion path
    MotionPathLeft,
    /// Right motion path
    MotionPathRight,
    /// Up motion path
    MotionPathUp,
    /// Down motion path
    MotionPathDown,
    /// Diagonal up right
    MotionPathDiagonalUpRight,
    /// Diagonal down right
    MotionPathDiagonalDownRight,
    /// Arc down
    MotionPathArcDown,
    /// Arc up
    MotionPathArcUp,
    /// Circle
    MotionPathCircle,
    /// Diamond motion path
    MotionPathDiamond,
    /// Heart
    MotionPathHeart,
    /// Hexagon
    MotionPathHexagon,
    /// Octagon
    MotionPathOctagon,
    /// Pentagon
    MotionPathPentagon,
    /// Square
    MotionPathSquare,
    /// Star 4
    MotionPathStar4,
    /// Star 5
    MotionPathStar5,
    /// Star 6
    MotionPathStar6,
    /// Star 8
    MotionPathStar8,
    /// Triangle
    MotionPathTriangle,
    /// Loop de loop
    MotionPathLoopDeLoop,
    /// Curved X
    MotionPathCurvedX,
    /// S curve 1
    MotionPathSCurve1,
    /// S curve 2
    MotionPathSCurve2,
    /// Sine wave
    MotionPathSineWave,
    /// Spiral left
    MotionPathSpiralLeft,
    /// Spiral right
    MotionPathSpiralRight,
    /// Spring
    MotionPathSpring,
    /// Zigzag
    MotionPathZigzag,

    /// Custom or unknown effect
    Custom,
}

/// Effect speed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum EffectSpeed {
    /// Very slow (5 seconds)
    VerySlow,
    /// Slow (3 seconds)
    Slow,
    /// Medium (2 seconds)
    #[default]
    Medium,
    /// Fast (1 second)
    Fast,
    /// Very fast (0.5 seconds)
    VeryFast,
}

impl EffectSpeed {
    /// Get duration in milliseconds.
    #[must_use]
    pub fn duration_ms(&self) -> u32 {
        match self {
            EffectSpeed::VerySlow => 5000,
            EffectSpeed::Slow => 3000,
            EffectSpeed::Medium => 2000,
            EffectSpeed::Fast => 1000,
            EffectSpeed::VeryFast => 500,
        }
    }
}

/// Effect direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum EffectDirection {
    /// No direction
    #[default]
    None,
    /// From top
    FromTop,
    /// From bottom
    FromBottom,
    /// From left
    FromLeft,
    /// From right
    FromRight,
    /// From top-left
    FromTopLeft,
    /// From top-right
    FromTopRight,
    /// From bottom-left
    FromBottomLeft,
    /// From bottom-right
    FromBottomRight,
    /// Horizontal
    Horizontal,
    /// Vertical
    Vertical,
    /// In (toward center)
    In,
    /// Out (from center)
    Out,
    /// Across
    Across,
    /// Clockwise
    Clockwise,
    /// Counter-clockwise
    CounterClockwise,
}

/// After-effect behavior for animations.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum AfterEffect {
    /// No after-effect
    #[default]
    None,
    /// Dim to color after animation
    DimToColor,
    /// Hide after animation
    Hide,
    /// Hide on next mouse click
    HideOnNextClick,
}

/// Animation trigger.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum AnimationTrigger {
    /// On click
    #[default]
    OnClick,
    /// With previous
    WithPrevious,
    /// After previous
    AfterPrevious,
}

/// Time node container for advanced animation timeline.
#[derive(Debug, Clone)]
pub struct TimeNodeContainer {
    /// Node type
    pub node_type: TimeNodeType,
    /// Duration in milliseconds
    pub duration: Option<u32>,
    /// Delay before start in milliseconds
    pub delay: u32,
    /// Fill mode (what happens after animation)
    pub fill: FillMode,
    /// Restart mode
    pub restart: RestartMode,
    /// Child nodes
    pub children: Vec<TimeNodeContainer>,
    /// Raw record for advanced parsing
    pub raw_record: Option<Record>,
}

impl Default for TimeNodeContainer {
    fn default() -> Self {
        Self {
            node_type: TimeNodeType::Sequence,
            duration: None,
            delay: 0,
            fill: FillMode::Hold,
            restart: RestartMode::Never,
            children: Vec::new(),
            raw_record: None,
        }
    }
}

/// Time node type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TimeNodeType {
    /// Parallel (children run simultaneously)
    Parallel,
    /// Sequence (children run one after another)
    #[default]
    Sequence,
    /// Effect (leaf node with actual effect)
    Effect,
    /// Audio
    Audio,
    /// Video
    Video,
}

/// Fill mode (what happens after animation).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum FillMode {
    /// Remove (hide after animation)
    Remove,
    /// Freeze (keep last frame)
    Freeze,
    /// Hold (same as freeze)
    #[default]
    Hold,
    /// Transition (fade to final state)
    Transition,
}

/// Restart mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum RestartMode {
    /// Always restart
    Always,
    /// When not active
    WhenNotActive,
    /// Never restart
    #[default]
    Never,
}
