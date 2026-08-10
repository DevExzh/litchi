//! Slide transition data types.
//!
//! Provides structures representing `PowerPoint` slide transitions.

/// Slide transition information.
#[derive(Debug, Clone, PartialEq)]
pub struct TransitionInfo {
    /// Transition type
    pub transition_type: TransitionType,
    /// Transition speed
    pub speed: TransitionSpeed,
    /// Advance mode (on click, automatic, or both)
    pub advance_mode: AdvanceMode,
    /// Automatic advance time in milliseconds (if `advance_mode` includes automatic)
    pub advance_time_ms: Option<u32>,
    /// Transition direction
    pub direction: TransitionDirection,
    /// Sound action
    pub sound: Option<SoundAction>,
    /// Loop sound until next sound
    pub loop_sound: bool,
}

impl Default for TransitionInfo {
    fn default() -> Self {
        Self::new()
    }
}

impl TransitionInfo {
    /// Create a new transition with no effect.
    #[must_use]
    pub fn new() -> Self {
        Self {
            transition_type: TransitionType::None,
            speed: TransitionSpeed::Medium,
            advance_mode: AdvanceMode::OnClick,
            advance_time_ms: None,
            direction: TransitionDirection::None,
            sound: None,
            loop_sound: false,
        }
    }

    /// Create a transition with a specific type.
    #[must_use]
    pub fn with_type(transition_type: TransitionType) -> Self {
        Self {
            transition_type,
            ..Self::new()
        }
    }

    /// Set transition speed.
    #[must_use]
    pub fn with_speed(mut self, speed: TransitionSpeed) -> Self {
        self.speed = speed;
        self
    }

    /// Set advance mode.
    #[must_use]
    pub fn with_advance_mode(mut self, advance_mode: AdvanceMode) -> Self {
        self.advance_mode = advance_mode;
        self
    }

    /// Set automatic advance time in milliseconds.
    #[must_use]
    pub fn with_advance_time(mut self, time_ms: u32) -> Self {
        self.advance_time_ms = Some(time_ms);
        self
    }

    /// Set transition direction.
    #[must_use]
    pub fn with_direction(mut self, direction: TransitionDirection) -> Self {
        self.direction = direction;
        self
    }

    /// Check if this transition has an effect.
    #[must_use]
    pub fn has_effect(&self) -> bool {
        self.transition_type != TransitionType::None
    }
}

/// Slide transition type.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TransitionType {
    /// No visible transition (`Cut` without the through-black variant).
    #[default]
    None,
    /// Blinds (horizontal or vertical)
    Blinds,
    /// Box in or out
    Box,
    /// Checkerboard (across or down)
    Checkerboard,
    /// Cover (left, right, up, down)
    Cover,
    /// Cut (through black)
    Cut,
    /// Dissolve
    Dissolve,
    /// Fade (through black or smoothly)
    Fade,
    /// Random bars (horizontal or vertical)
    RandomBars,
    /// Split (horizontal or vertical, in or out)
    Split,
    /// Strips (left-down, left-up, right-down, right-up)
    Strips,
    /// Uncover (left, right, up, down)
    Uncover,
    /// Wipe (left, right, up, down)
    Wipe,
    /// Push (left, right, up, down)
    Push,
    /// Comb (horizontal or vertical)
    Comb,
    /// Wheel (1-8 spokes)
    Wheel,
    /// Wedge
    Wedge,
    /// Zoom (not representable in `[MS-PPT]` 2.6.6).
    Zoom,
    /// Random
    Random,
    /// Newsflash
    Newsflash,
    /// Diamond.
    Diamond,
    /// Plus.
    Plus,
    /// Alpha fade.
    AlphaFade,
    /// Circle.
    Circle,
    /// Undefined effect (`effectType = 255`), which consumers must ignore.
    Undefined,
    /// Vortex (not representable in `[MS-PPT]` 2.6.6).
    Vortex,
    /// Shred (not representable in `[MS-PPT]` 2.6.6).
    Shred,
    /// Switch (not representable in `[MS-PPT]` 2.6.6).
    Switch,
    /// Flip (not representable in `[MS-PPT]` 2.6.6).
    Flip,
    /// Gallery (not representable in `[MS-PPT]` 2.6.6).
    Gallery,
    /// Cube (not representable in `[MS-PPT]` 2.6.6).
    Cube,
    /// Doors (not representable in `[MS-PPT]` 2.6.6).
    Doors,
    /// Window (not representable in `[MS-PPT]` 2.6.6).
    Window,
    /// Ferris (not representable in `[MS-PPT]` 2.6.6).
    Ferris,
    /// Conveyor (not representable in `[MS-PPT]` 2.6.6).
    Conveyor,
    /// Rotate (not representable in `[MS-PPT]` 2.6.6).
    Rotate,
    /// Pan (not representable in `[MS-PPT]` 2.6.6).
    Pan,
    /// Glitter (not representable in `[MS-PPT]` 2.6.6).
    Glitter,
    /// Honeycomb (not representable in `[MS-PPT]` 2.6.6).
    Honeycomb,
    /// Flash (not representable in `[MS-PPT]` 2.6.6).
    Flash,
    /// Ripple (not representable in `[MS-PPT]` 2.6.6).
    Ripple,
    /// Fracture (not representable in `[MS-PPT]` 2.6.6).
    Fracture,
    /// Crush (not representable in `[MS-PPT]` 2.6.6).
    Crush,
    /// Peel (not representable in `[MS-PPT]` 2.6.6).
    Peel,
    /// `PageCurl` (not representable in `[MS-PPT]` 2.6.6).
    PageCurl,
    /// Airplane (not representable in `[MS-PPT]` 2.6.6).
    Airplane,
    /// Origami (not representable in `[MS-PPT]` 2.6.6).
    Origami,
    /// Morph (not representable in `[MS-PPT]` 2.6.6).
    Morph,
}

/// Transition speed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TransitionSpeed {
    /// Slow (0.75 seconds).
    Slow,
    /// Medium (0.5 seconds).
    #[default]
    Medium,
    /// Fast (0.25 seconds).
    Fast,
}

impl TransitionSpeed {
    /// Get duration in milliseconds.
    #[must_use]
    pub fn duration_ms(&self) -> u32 {
        match self {
            TransitionSpeed::Slow => 750,
            TransitionSpeed::Medium => 500,
            TransitionSpeed::Fast => 250,
        }
    }
}

/// Slide advance mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum AdvanceMode {
    /// Advance on mouse click only
    #[default]
    OnClick,
    /// Advance automatically after time
    Automatic,
    /// Advance on click or automatically (both)
    Both,
}

/// Transition direction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TransitionDirection {
    /// No specific direction
    #[default]
    None,
    /// Horizontal
    Horizontal,
    /// Vertical
    Vertical,
    /// From left
    FromLeft,
    /// From right
    FromRight,
    /// From top
    FromTop,
    /// From bottom
    FromBottom,
    /// In (toward center)
    In,
    /// Out (from center)
    Out,
    /// Left-down
    LeftDown,
    /// Left-up
    LeftUp,
    /// Right-down
    RightDown,
    /// Right-up
    RightUp,
    /// Cut through black.
    ThroughBlack,
    /// Split horizontally out.
    HorizontalOut,
    /// Split horizontally in.
    HorizontalIn,
    /// Split vertically out.
    VerticalOut,
    /// Split vertically in.
    VerticalIn,
    /// Wheel with one radial division.
    Spokes1,
    /// Wheel with two radial divisions.
    Spokes2,
    /// Wheel with three radial divisions.
    Spokes3,
    /// Wheel with four radial divisions.
    Spokes4,
    /// Wheel with eight radial divisions.
    Spokes8,
}

/// Sound action for transition.
#[derive(Debug, Clone, PartialEq)]
pub struct SoundAction {
    /// Sound name or identifier
    pub name: String,
    /// Whether sound is built-in
    pub is_builtin: bool,
    /// Sound file reference (for external sounds)
    pub file_ref: Option<String>,
}

impl SoundAction {
    /// Create a built-in sound action.
    pub fn builtin(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            is_builtin: true,
            file_ref: None,
        }
    }

    /// Create an external sound action.
    pub fn external(name: impl Into<String>, file_ref: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            is_builtin: false,
            file_ref: Some(file_ref.into()),
        }
    }
}
