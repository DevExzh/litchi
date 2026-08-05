//! Archive-free semantic values for Keynote build animations.
//!
//! This module deliberately contains no package, archive, protobuf, object, or
//! runtime identity.  The IWA adapter can decode native fields into these
//! values and publish them only after validation.  Native discriminants that
//! still need lossless preservation remain behind the existing adapter seam;
//! ordinary build settings use the typed values below instead of raw fields.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::float_cmp,
    clippy::missing_errors_doc,
    clippy::shadow_reuse,
    clippy::unnecessary_lazy_evaluations,
    clippy::wildcard_enum_match_arm,
    reason = "This focused migration leaf keeps its validated constructors and lossless unknown branches together; the archive adapter audit will replace these local exceptions as each native family moves."
)]

use super::{Acceleration, Start};
use crate::{Error, Result, Seconds};

/// Maximum UTF-8 byte length of one build identifier or unknown value.
pub const MAX_IDENTIFIER_BYTES: usize = 64 * 1024;
/// Maximum number of continuous subpaths in one motion path.
pub const MAX_PATH_SUBPATHS: usize = 256;
/// Maximum number of nodes in one motion path.
pub const MAX_PATH_NODES: usize = 16 * 1024;

const APPEAR_EFFECT: &str = "appear";
const DISSOLVE_EFFECT: &str = "dissolve";
const MOVE_IN_EFFECT: &str = "move";
const SCALE_EFFECT: &str = "scale";
const FADE_AND_SCALE_EFFECT: &str = "fade-scale";
const ROTATE_ACTION_EFFECT: &str = "apple:action-rotation";
const SCALE_ACTION_EFFECT: &str = "apple:action-scale";
const OPACITY_ACTION_EFFECT: &str = "apple:action-opacity";
const MOTION_ACTION_EFFECT: &str = "apple:action-motion-path";
const BLINK_ACTION_EFFECT: &str = "apple:action-blink";
const BOUNCE_ACTION_EFFECT: &str = "apple:action-bounce";
const FLIP_ACTION_EFFECT: &str = "apple:action-flip";
const JIGGLE_ACTION_EFFECT: &str = "apple:action-jiggle";
const POP_ACTION_EFFECT: &str = "apple:action-pop";
const PULSE_ACTION_EFFECT: &str = "apple:action-pulse";
const KEYBOARD_EFFECT: &str = "apple:keyboard";
const OBJECT_DISSOLVE_EFFECT: &str = "apple:dissolve character";
const OBJECT_SHIMMER_EFFECT: &str = "com.apple.iWork.Keynote.KLNShimmer";
const OBJECT_SKID_EFFECT: &str = "com.apple.iWork.Keynote.KNBuildSkidByCharacter";
const OBJECT_SWOOSH_EFFECT: &str = "com.apple.iWork.Keynote.BLTSwoosh";
const OBJECT_TRACE_EFFECT: &str = "com.apple.iWork.Keynote.Trace";

fn validate_text(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::EmptyIdentifier);
    }
    if value.len() > MAX_IDENTIFIER_BYTES {
        return Err(Error::IdentifierTooLarge);
    }
    if value.contains('\0') {
        return Err(Error::NulString);
    }
    Ok(())
}

/// A bounded, NUL-free producer value that has no known semantic variant.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct UnknownText(Box<str>);

impl UnknownText {
    /// Construct a bounded unknown value.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let value = value.into();
        validate_text(&value)?;
        Ok(Self(value))
    }

    /// Borrow the preserved producer value.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for UnknownText {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// A finite scalar used by checked build geometry and effect parameters.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Finite(f64);

impl Finite {
    /// Construct a finite scalar.
    pub fn new(value: f64) -> Result<Self> {
        value
            .is_finite()
            .then_some(Self(value))
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the scalar value.
    #[must_use]
    pub const fn get(self) -> f64 {
        self.0
    }
}

/// A typed build effect.
///
/// Parameterized effects carry their validated semantic parameters directly,
/// which makes impossible combinations harder to construct than the native
/// collection of unrelated optional fields.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Effect {
    /// An object appears.
    Appear,
    /// An object dissolves into view.
    Dissolve,
    /// An object moves into view.
    MoveIn,
    /// An object scales into view.
    Scale,
    /// An object fades and scales into view.
    FadeAndScale,
    /// A typed action effect.
    Action(Action),
    /// A typed emphasis effect.
    Emphasis(Emphasis),
    /// A typed keyboard build effect.
    Keyboard(Keyboard),
    /// A typed object build effect.
    Object(ObjectEffect),
    /// A producer effect not yet modeled by this crate.
    Unknown(UnknownText),
}

impl Effect {
    /// Decode a simple producer identifier without allocating for known values.
    ///
    /// Parameterized native effects are retained as bounded unknown text until
    /// their native parameters have been decoded into one of the typed
    /// constructors on this module.
    pub fn from_identifier(identifier: &str) -> Result<Self> {
        validate_text(identifier)?;
        let effect = if identifier.eq_ignore_ascii_case(APPEAR_EFFECT) {
            Self::Appear
        } else if identifier.eq_ignore_ascii_case(DISSOLVE_EFFECT) {
            Self::Dissolve
        } else if identifier.eq_ignore_ascii_case(MOVE_IN_EFFECT)
            || contains_ascii_case_insensitive(identifier, b"move")
        {
            Self::MoveIn
        } else if identifier.eq_ignore_ascii_case(FADE_AND_SCALE_EFFECT)
            || (contains_ascii_case_insensitive(identifier, b"fade")
                && contains_ascii_case_insensitive(identifier, b"scale"))
        {
            Self::FadeAndScale
        } else if identifier.eq_ignore_ascii_case(SCALE_EFFECT)
            || contains_ascii_case_insensitive(identifier, b"scale")
        {
            Self::Scale
        } else {
            Self::Unknown(UnknownText(identifier.into()))
        };
        Ok(effect)
    }

    /// Construct an unknown effect for a value that is not a named simple
    /// effect.
    pub fn unknown(value: impl Into<Box<str>>) -> Result<Self> {
        let value = UnknownText::new(value)?;
        if is_simple_effect_identifier(value.as_str()) {
            return Err(Error::NonCanonicalEffect);
        }
        Ok(Self::Unknown(value))
    }

    /// Construct a parameterized action effect.
    #[must_use]
    pub const fn action(action: Action) -> Self {
        Self::Action(action)
    }

    /// Construct a parameterized emphasis effect.
    #[must_use]
    pub const fn emphasis(emphasis: Emphasis) -> Self {
        Self::Emphasis(emphasis)
    }

    /// Construct a parameterized keyboard effect.
    #[must_use]
    pub const fn keyboard(keyboard: Keyboard) -> Self {
        Self::Keyboard(keyboard)
    }

    /// Construct a parameterized object effect.
    #[must_use]
    pub const fn object(effect: ObjectEffect) -> Self {
        Self::Object(effect)
    }

    /// Return the canonical producer identifier or preserved unknown text.
    #[must_use]
    pub fn identifier(&self) -> &str {
        match self {
            Self::Appear => APPEAR_EFFECT,
            Self::Dissolve => DISSOLVE_EFFECT,
            Self::MoveIn => MOVE_IN_EFFECT,
            Self::Scale => SCALE_EFFECT,
            Self::FadeAndScale => FADE_AND_SCALE_EFFECT,
            Self::Action(action) => action.identifier(),
            Self::Emphasis(emphasis) => emphasis.identifier(),
            Self::Keyboard(_) => KEYBOARD_EFFECT,
            Self::Object(effect) => effect.identifier(),
            Self::Unknown(value) => value.as_str(),
        }
    }

    /// Validate the value before a native adapter publishes it.
    pub fn validate(&self) -> Result<()> {
        match self {
            Self::Action(action) => action.validate(),
            Self::Emphasis(emphasis) => emphasis.validate(),
            Self::Unknown(value) if is_simple_effect_identifier(value.as_str()) => {
                Err(Error::NonCanonicalEffect)
            },
            _ => Ok(()),
        }
    }
}

/// The action family of build effects.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Action {
    /// Rotate an object by a checked number of degrees.
    Rotate(Rotation),
    /// Scale an object by a checked positive factor.
    Scale(Scale),
    /// Set an object's final opacity percentage.
    Opacity(Opacity),
    /// Move an object along a checked path.
    Move(Motion),
    /// A future action whose parameters are not modeled yet.
    Unknown(UnknownText),
}

impl Action {
    /// Construct an unknown action identifier.
    pub fn unknown(value: impl Into<Box<str>>) -> Result<Self> {
        let value = UnknownText::new(value)?;
        if is_action_identifier(value.as_str()) {
            return Err(Error::NonCanonicalEffect);
        }
        Ok(Self::Unknown(value))
    }

    /// Return the producer identifier for this action.
    #[must_use]
    pub fn identifier(&self) -> &str {
        match self {
            Self::Rotate(_) => ROTATE_ACTION_EFFECT,
            Self::Scale(_) => SCALE_ACTION_EFFECT,
            Self::Opacity(_) => OPACITY_ACTION_EFFECT,
            Self::Move(_) => MOTION_ACTION_EFFECT,
            Self::Unknown(value) => value.as_str(),
        }
    }

    fn validate(&self) -> Result<()> {
        match self {
            Self::Rotate(value) => value.validate(),
            Self::Scale(value) => value.validate(),
            Self::Opacity(value) => value.validate(),
            Self::Move(value) => value.validate(),
            Self::Unknown(value) if is_action_identifier(value.as_str()) => {
                Err(Error::NonCanonicalEffect)
            },
            _ => Ok(()),
        }
    }
}

/// Rotation direction for a typed action.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RotationDirection {
    /// Rotate clockwise.
    Clockwise,
    /// Rotate counterclockwise.
    Counterclockwise,
}

/// Checked rotation action parameters.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rotation {
    total_degrees: Finite,
    direction: RotationDirection,
    acceleration: Acceleration,
}

impl Rotation {
    /// Construct rotation parameters with a positive finite angle.
    pub fn new(
        total_degrees: f64,
        direction: RotationDirection,
        acceleration: Acceleration,
    ) -> Result<Self> {
        let total_degrees = Finite::new(total_degrees)?;
        (total_degrees.get() > 0.0)
            .then_some(Self {
                total_degrees,
                direction,
                acceleration,
            })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the total rotation in degrees.
    #[must_use]
    pub const fn total_degrees(self) -> f64 {
        self.total_degrees.get()
    }

    /// Return the rotation direction.
    #[must_use]
    pub const fn direction(self) -> RotationDirection {
        self.direction
    }

    /// Return the timing acceleration.
    #[must_use]
    pub const fn acceleration(self) -> Acceleration {
        self.acceleration
    }

    fn validate(self) -> Result<()> {
        (self.total_degrees.get() > 0.0)
            .then_some(())
            .ok_or(Error::InvalidBuildValue)
    }
}

/// Checked scale action parameters.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Scale {
    factor: Finite,
    acceleration: Acceleration,
}

impl Scale {
    /// Construct a scale action with a positive finite factor.
    pub fn new(factor: f64, acceleration: Acceleration) -> Result<Self> {
        let factor = Finite::new(factor)?;
        (factor.get() > 0.0)
            .then_some(Self {
                factor,
                acceleration,
            })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the final scale factor.
    #[must_use]
    pub const fn factor(self) -> f64 {
        self.factor.get()
    }

    /// Return the timing acceleration.
    #[must_use]
    pub const fn acceleration(self) -> Acceleration {
        self.acceleration
    }

    fn validate(self) -> Result<()> {
        (self.factor.get() > 0.0)
            .then_some(())
            .ok_or(Error::InvalidBuildValue)
    }
}

/// Checked opacity action parameters.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Opacity {
    percent: Finite,
    acceleration: Acceleration,
}

impl Opacity {
    /// Construct an opacity action in the inclusive 0–100 percent range.
    pub fn new(percent: f64, acceleration: Acceleration) -> Result<Self> {
        let percent = Finite::new(percent)?;
        (0.0..=100.0)
            .contains(&percent.get())
            .then_some(Self {
                percent,
                acceleration,
            })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the final opacity percentage.
    #[must_use]
    pub const fn percent(self) -> f64 {
        self.percent.get()
    }

    /// Return the timing acceleration.
    #[must_use]
    pub const fn acceleration(self) -> Acceleration {
        self.acceleration
    }

    fn validate(self) -> Result<()> {
        (0.0..=100.0)
            .contains(&self.percent.get())
            .then_some(())
            .ok_or(Error::InvalidBuildValue)
    }
}

/// Checked motion-path action parameters.
#[derive(Debug, Clone, PartialEq)]
pub struct Motion {
    path: Path,
    align_to_path: bool,
    acceleration: Acceleration,
}

impl Motion {
    /// Construct a motion action from a validated path.
    #[must_use]
    pub const fn new(path: Path, align_to_path: bool, acceleration: Acceleration) -> Self {
        Self {
            path,
            align_to_path,
            acceleration,
        }
    }

    /// Borrow the motion path.
    #[must_use]
    pub const fn path(&self) -> &Path {
        &self.path
    }

    /// Return whether the object follows the path tangent.
    #[must_use]
    pub const fn align_to_path(&self) -> bool {
        self.align_to_path
    }

    /// Return the timing acceleration.
    #[must_use]
    pub const fn acceleration(&self) -> Acceleration {
        self.acceleration
    }

    fn validate(&self) -> Result<()> {
        self.path.validate()
    }
}

/// Horizontal direction used by flip emphasis and object effects.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum HorizontalDirection {
    /// Traverse from left to right.
    LeftToRight,
    /// Traverse from right to left.
    RightToLeft,
}

/// Origin used by a Swoosh object effect.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SwooshDirection {
    /// Start at the center.
    Center,
    /// Start at the left edge.
    FromLeft,
    /// Start at the right edge.
    FromRight,
}

/// Text traversal direction for a Keyboard build.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeyboardDirection {
    /// Reveal text from the beginning.
    Forward,
    /// Reveal text from the end.
    Backward,
}

/// Checked parameters for a Keyboard build.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Keyboard {
    direction: KeyboardDirection,
    show_cursor: bool,
}

impl Keyboard {
    /// Construct Keyboard build parameters.
    #[must_use]
    pub const fn new(direction: KeyboardDirection, show_cursor: bool) -> Self {
        Self {
            direction,
            show_cursor,
        }
    }

    /// Return the text traversal direction.
    #[must_use]
    pub const fn direction(self) -> KeyboardDirection {
        self.direction
    }

    /// Return whether Keynote displays the insertion cursor.
    #[must_use]
    pub const fn show_cursor(self) -> bool {
        self.show_cursor
    }
}

/// Typed object build effects with no opaque parameters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ObjectEffect {
    /// Dissolve by character.
    Dissolve,
    /// Shimmer into view.
    Shimmer,
    /// Skid horizontally.
    Skid(HorizontalDirection),
    /// Swoosh from a semantic origin.
    Swoosh(SwooshDirection),
    /// Trace horizontally.
    Trace(HorizontalDirection),
}

impl ObjectEffect {
    /// Return the native identifier represented by the typed effect.
    #[must_use]
    pub const fn identifier(self) -> &'static str {
        match self {
            Self::Dissolve => OBJECT_DISSOLVE_EFFECT,
            Self::Shimmer => OBJECT_SHIMMER_EFFECT,
            Self::Skid(_) => OBJECT_SKID_EFFECT,
            Self::Swoosh(_) => OBJECT_SWOOSH_EFFECT,
            Self::Trace(_) => OBJECT_TRACE_EFFECT,
        }
    }
}

/// Emphasis action intensity used by Jiggle.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum JiggleIntensity {
    /// Small jiggle.
    Small,
    /// Medium jiggle.
    Medium,
    /// Large jiggle.
    Large,
}

/// Direction used by Flip emphasis.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FlipDirection {
    /// Flip from left to right.
    LeftToRight,
    /// Flip from right to left.
    RightToLeft,
}

/// Checked Blink emphasis parameters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Blink {
    repeat_count: u32,
    fade: bool,
}

impl Blink {
    /// Construct Blink parameters with a positive repeat count.
    pub fn new(repeat_count: u32, fade: bool) -> Result<Self> {
        (repeat_count > 0)
            .then_some(Self { repeat_count, fade })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the repeat count.
    #[must_use]
    pub const fn repeat_count(self) -> u32 {
        self.repeat_count
    }

    /// Return whether the emphasis fades between blinks.
    #[must_use]
    pub const fn fade(self) -> bool {
        self.fade
    }
}

/// Checked Bounce emphasis parameters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Bounce {
    repeat_count: u32,
    decay: bool,
}

impl Bounce {
    /// Construct Bounce parameters with a positive repeat count.
    pub fn new(repeat_count: u32, decay: bool) -> Result<Self> {
        (repeat_count > 0)
            .then_some(Self {
                repeat_count,
                decay,
            })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the repeat count.
    #[must_use]
    pub const fn repeat_count(self) -> u32 {
        self.repeat_count
    }

    /// Return whether the bounce decays.
    #[must_use]
    pub const fn decay(self) -> bool {
        self.decay
    }
}

/// Checked Flip emphasis parameters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Flip {
    repeat_count: u32,
    direction: FlipDirection,
}

impl Flip {
    /// Construct Flip parameters with a positive repeat count.
    pub fn new(repeat_count: u32, direction: FlipDirection) -> Result<Self> {
        (repeat_count > 0)
            .then_some(Self {
                repeat_count,
                direction,
            })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the repeat count.
    #[must_use]
    pub const fn repeat_count(self) -> u32 {
        self.repeat_count
    }

    /// Return the flip direction.
    #[must_use]
    pub const fn direction(self) -> FlipDirection {
        self.direction
    }
}

/// Checked Jiggle emphasis parameters.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Jiggle {
    intensity: JiggleIntensity,
}

impl Jiggle {
    /// Construct Jiggle parameters.
    #[must_use]
    pub const fn new(intensity: JiggleIntensity) -> Self {
        Self { intensity }
    }

    /// Return the jiggle intensity.
    #[must_use]
    pub const fn intensity(self) -> JiggleIntensity {
        self.intensity
    }
}

/// Checked Pop emphasis parameters.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Pop {
    scale_percent: Finite,
}

impl Pop {
    /// Construct Pop parameters with a positive finite scale percentage.
    pub fn new(scale_percent: f64) -> Result<Self> {
        let scale_percent = Finite::new(scale_percent)?;
        (scale_percent.get() > 0.0)
            .then_some(Self { scale_percent })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the scale percentage.
    #[must_use]
    pub const fn scale_percent(self) -> f64 {
        self.scale_percent.get()
    }
}

/// Checked Pulse emphasis parameters.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Pulse {
    repeat_count: u32,
    scale_percent: Finite,
}

impl Pulse {
    /// Construct Pulse parameters with a positive repeat count and scale.
    pub fn new(repeat_count: u32, scale_percent: f64) -> Result<Self> {
        let scale_percent = Finite::new(scale_percent)?;
        (repeat_count > 0 && scale_percent.get() > 0.0)
            .then_some(Self {
                repeat_count,
                scale_percent,
            })
            .ok_or(Error::InvalidBuildValue)
    }

    /// Return the repeat count.
    #[must_use]
    pub const fn repeat_count(self) -> u32 {
        self.repeat_count
    }

    /// Return the scale percentage.
    #[must_use]
    pub const fn scale_percent(self) -> f64 {
        self.scale_percent.get()
    }
}

/// Typed emphasis action parameters.
#[derive(Debug, Clone, PartialEq)]
#[non_exhaustive]
pub enum Emphasis {
    /// Blink repeatedly.
    Blink(Blink),
    /// Bounce repeatedly.
    Bounce(Bounce),
    /// Flip repeatedly.
    Flip(Flip),
    /// Jiggle at a checked intensity.
    Jiggle(Jiggle),
    /// Pop by a checked scale percentage.
    Pop(Pop),
    /// Pulse repeatedly by a checked scale percentage.
    Pulse(Pulse),
    /// A future emphasis action.
    Unknown(UnknownText),
}

impl Emphasis {
    /// Construct Blink emphasis parameters.
    pub fn blink(repeat_count: u32, fade: bool) -> Result<Self> {
        Ok(Self::Blink(Blink::new(repeat_count, fade)?))
    }

    /// Construct Bounce emphasis parameters.
    pub fn bounce(repeat_count: u32, decay: bool) -> Result<Self> {
        Ok(Self::Bounce(Bounce::new(repeat_count, decay)?))
    }

    /// Construct Flip emphasis parameters.
    pub fn flip(repeat_count: u32, direction: FlipDirection) -> Result<Self> {
        Ok(Self::Flip(Flip::new(repeat_count, direction)?))
    }

    /// Construct Jiggle emphasis parameters.
    #[must_use]
    pub const fn jiggle(intensity: JiggleIntensity) -> Self {
        Self::Jiggle(Jiggle::new(intensity))
    }

    /// Construct Pop emphasis parameters.
    pub fn pop(scale_percent: f64) -> Result<Self> {
        Ok(Self::Pop(Pop::new(scale_percent)?))
    }

    /// Construct Pulse emphasis parameters.
    pub fn pulse(repeat_count: u32, scale_percent: f64) -> Result<Self> {
        Ok(Self::Pulse(Pulse::new(repeat_count, scale_percent)?))
    }

    /// Construct a future emphasis action while preserving its identifier.
    pub fn unknown(value: impl Into<Box<str>>) -> Result<Self> {
        let value = UnknownText::new(value)?;
        if is_emphasis_identifier(value.as_str()) {
            return Err(Error::NonCanonicalEffect);
        }
        Ok(Self::Unknown(value))
    }

    /// Return the producer identifier for this emphasis action.
    #[must_use]
    pub fn identifier(&self) -> &str {
        match self {
            Self::Blink(_) => BLINK_ACTION_EFFECT,
            Self::Bounce(_) => BOUNCE_ACTION_EFFECT,
            Self::Flip(_) => FLIP_ACTION_EFFECT,
            Self::Jiggle(_) => JIGGLE_ACTION_EFFECT,
            Self::Pop(_) => POP_ACTION_EFFECT,
            Self::Pulse(_) => PULSE_ACTION_EFFECT,
            Self::Unknown(value) => value.as_str(),
        }
    }

    fn validate(&self) -> Result<()> {
        let valid = match self {
            Self::Blink(value) => value.repeat_count > 0,
            Self::Bounce(value) => value.repeat_count > 0,
            Self::Flip(value) => value.repeat_count > 0,
            Self::Jiggle(_) => true,
            Self::Pop(value) => value.scale_percent.get() > 0.0,
            Self::Pulse(value) => value.repeat_count > 0 && value.scale_percent.get() > 0.0,
            Self::Unknown(value) => !is_emphasis_identifier(value.as_str()),
        };
        valid.then_some(()).ok_or_else(|| {
            if matches!(self, Self::Unknown(_)) {
                Error::NonCanonicalEffect
            } else {
                Error::InvalidBuildValue
            }
        })
    }
}

/// Node interpolation kind for a motion path.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub enum NodeKind {
    /// A sharp corner.
    Sharp,
    /// A Bézier node.
    Bezier,
    /// A smooth node.
    Smooth,
}

/// A checked finite point in build-path coordinates.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Point {
    x: Finite,
    y: Finite,
}

impl Point {
    /// Construct a point with finite coordinates.
    pub fn new(x: f64, y: f64) -> Result<Self> {
        Ok(Self {
            x: Finite::new(x)?,
            y: Finite::new(y)?,
        })
    }

    /// Return the horizontal coordinate.
    #[must_use]
    pub const fn x(self) -> f64 {
        self.x.get()
    }

    /// Return the vertical coordinate.
    #[must_use]
    pub const fn y(self) -> f64 {
        self.y.get()
    }
}

/// One checked motion-path node.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Node {
    in_control_point: Point,
    point: Point,
    out_control_point: Point,
    kind: NodeKind,
}

impl Node {
    /// Construct a node from its checked control points.
    #[must_use]
    pub const fn new(
        in_control_point: Point,
        point: Point,
        out_control_point: Point,
        kind: NodeKind,
    ) -> Self {
        Self {
            in_control_point,
            point,
            out_control_point,
            kind,
        }
    }

    /// Construct a sharp node whose control points coincide with its point.
    #[must_use]
    pub const fn sharp(point: Point) -> Self {
        Self::new(point, point, point, NodeKind::Sharp)
    }

    /// Return the incoming control point.
    #[must_use]
    pub const fn in_control_point(self) -> Point {
        self.in_control_point
    }

    /// Return the node point.
    #[must_use]
    pub const fn point(self) -> Point {
        self.point
    }

    /// Return the outgoing control point.
    #[must_use]
    pub const fn out_control_point(self) -> Point {
        self.out_control_point
    }

    /// Return the interpolation kind.
    #[must_use]
    pub const fn kind(self) -> NodeKind {
        self.kind
    }
}

/// A continuous path segment stored in exact-size boxed node storage.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct Subpath {
    nodes: Box<[Node]>,
    closed: bool,
}

impl Subpath {
    /// Construct a subpath with at least two nodes.
    pub fn new(nodes: impl Into<Box<[Node]>>, closed: bool) -> Result<Self> {
        let nodes = nodes.into();
        if nodes.len() < 2 {
            return Err(Error::InvalidBuildPath);
        }
        if nodes.len() > MAX_PATH_NODES {
            return Err(Error::BuildTooLarge);
        }
        Ok(Self { nodes, closed })
    }

    /// Borrow the exact node sequence.
    #[must_use]
    pub fn nodes(&self) -> &[Node] {
        &self.nodes
    }

    /// Return whether the subpath is closed.
    #[must_use]
    pub const fn closed(&self) -> bool {
        self.closed
    }
}

/// A bounded editable motion path.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct Path {
    subpaths: Box<[Subpath]>,
    natural_width: Finite,
    natural_height: Finite,
    horizontal_flip: bool,
    vertical_flip: bool,
}

impl Path {
    /// Construct a path from bounded subpaths and non-negative natural size.
    pub fn new(
        subpaths: impl Into<Box<[Subpath]>>,
        natural_width: f64,
        natural_height: f64,
        horizontal_flip: bool,
        vertical_flip: bool,
    ) -> Result<Self> {
        let subpaths = subpaths.into();
        if subpaths.is_empty() {
            return Err(Error::InvalidBuildPath);
        }
        if subpaths.len() > MAX_PATH_SUBPATHS {
            return Err(Error::BuildTooLarge);
        }
        let natural_width = Finite::new(natural_width)?;
        let natural_height = Finite::new(natural_height)?;
        if natural_width.get() < 0.0
            || natural_height.get() < 0.0
            || (natural_width.get() == 0.0 && natural_height.get() == 0.0)
        {
            return Err(Error::InvalidBuildPath);
        }
        let node_count = subpaths
            .iter()
            .map(|subpath| subpath.nodes.len())
            .sum::<usize>();
        if node_count > MAX_PATH_NODES {
            return Err(Error::BuildTooLarge);
        }
        Ok(Self {
            subpaths,
            natural_width,
            natural_height,
            horizontal_flip,
            vertical_flip,
        })
    }

    /// Construct a two-node straight motion path.
    pub fn straight(delta_x: f64, delta_y: f64) -> Result<Self> {
        let origin = Point::new(0.0, 0.0)?;
        let destination = Point::new(delta_x, delta_y)?;
        let subpath = Subpath::new(
            vec![Node::sharp(origin), Node::sharp(destination)].into_boxed_slice(),
            false,
        )?;
        Self::new(
            vec![subpath].into_boxed_slice(),
            delta_x.abs(),
            delta_y.abs(),
            false,
            false,
        )
    }

    /// Borrow the bounded subpaths.
    #[must_use]
    pub fn subpaths(&self) -> &[Subpath] {
        &self.subpaths
    }

    /// Return the natural width in points.
    #[must_use]
    pub const fn natural_width(&self) -> f64 {
        self.natural_width.get()
    }

    /// Return the natural height in points.
    #[must_use]
    pub const fn natural_height(&self) -> f64 {
        self.natural_height.get()
    }

    /// Return the horizontal flip state.
    #[must_use]
    pub const fn horizontal_flip(&self) -> bool {
        self.horizontal_flip
    }

    /// Return the vertical flip state.
    #[must_use]
    pub const fn vertical_flip(&self) -> bool {
        self.vertical_flip
    }

    fn validate(&self) -> Result<()> {
        if self.subpaths.is_empty()
            || self.subpaths.len() > MAX_PATH_SUBPATHS
            || self.natural_width.get() < 0.0
            || self.natural_height.get() < 0.0
            || (self.natural_width.get() == 0.0 && self.natural_height.get() == 0.0)
        {
            return Err(Error::InvalidBuildPath);
        }
        let node_count = self
            .subpaths
            .iter()
            .map(|subpath| subpath.nodes.len())
            .sum::<usize>();
        (node_count <= MAX_PATH_NODES)
            .then_some(())
            .ok_or(Error::BuildTooLarge)
    }
}

/// A normalized timing curve backed by a bounded path.
#[derive(Debug, Clone, PartialEq, PartialOrd)]
pub struct TimingCurve {
    path: Path,
}

impl TimingCurve {
    /// Construct a linear timing curve without allocation beyond its boxed
    /// path storage.
    #[must_use]
    pub fn linear() -> Self {
        let origin = Point {
            x: Finite(0.0),
            y: Finite(0.0),
        };
        let destination = Point {
            x: Finite(1.0),
            y: Finite(1.0),
        };
        let path = Path {
            subpaths: vec![Subpath {
                nodes: vec![Node::sharp(origin), Node::sharp(destination)].into_boxed_slice(),
                closed: false,
            }]
            .into_boxed_slice(),
            natural_width: Finite(1.0),
            natural_height: Finite(1.0),
            horizontal_flip: false,
            vertical_flip: false,
        };
        Self { path }
    }

    /// Construct a one-segment cubic timing curve.
    pub fn cubic(first_control_point: Point, second_control_point: Point) -> Result<Self> {
        let origin = Point {
            x: Finite(0.0),
            y: Finite(0.0),
        };
        let destination = Point {
            x: Finite(1.0),
            y: Finite(1.0),
        };
        let subpath = Subpath::new(
            vec![
                Node::new(origin, origin, first_control_point, NodeKind::Bezier),
                Node::new(
                    second_control_point,
                    destination,
                    destination,
                    NodeKind::Bezier,
                ),
            ]
            .into_boxed_slice(),
            false,
        )?;
        Self::from_path(Path::new(
            vec![subpath].into_boxed_slice(),
            1.0,
            1.0,
            false,
            false,
        )?)
    }

    /// Construct a timing curve from a normalized path.
    pub fn from_path(path: Path) -> Result<Self> {
        path.validate()?;
        if path.horizontal_flip
            || path.vertical_flip
            || path.natural_width.get() != 1.0
            || path.natural_height.get() != 1.0
        {
            return Err(Error::InvalidBuildPath);
        }
        let first = path
            .subpaths
            .first()
            .and_then(|subpath| subpath.nodes.first())
            .ok_or(Error::InvalidBuildPath)?
            .point;
        let last = path
            .subpaths
            .last()
            .and_then(|subpath| subpath.nodes.last())
            .ok_or(Error::InvalidBuildPath)?
            .point;
        if first
            != (Point {
                x: Finite(0.0),
                y: Finite(0.0),
            })
            || last
                != (Point {
                    x: Finite(1.0),
                    y: Finite(1.0),
                })
        {
            return Err(Error::InvalidBuildPath);
        }
        if path
            .subpaths
            .iter()
            .flat_map(|subpath| subpath.nodes.iter())
            .flat_map(|node| [node.in_control_point, node.point, node.out_control_point])
            .any(|point| {
                !(0.0..=1.0).contains(&point.x.get()) || !(0.0..=1.0).contains(&point.y.get())
            })
        {
            return Err(Error::InvalidBuildPath);
        }
        Ok(Self { path })
    }

    /// Borrow the normalized path.
    #[must_use]
    pub const fn path(&self) -> &Path {
        &self.path
    }
}

/// Complete, validated settings for one build event.
#[derive(Debug, Clone, PartialEq)]
pub struct Settings {
    effect: Effect,
    start: Start,
    duration: Seconds,
    delay: Seconds,
}

impl Settings {
    /// Construct settings with an immediate start and no delay.
    #[must_use]
    pub const fn new(effect: Effect, duration: Seconds) -> Self {
        Self {
            effect,
            start: Start::OnClick,
            duration,
            delay: Seconds::ZERO,
        }
    }

    /// Return the typed build effect.
    #[must_use]
    pub const fn effect(&self) -> &Effect {
        &self.effect
    }

    /// Replace the effect after validating the candidate value.
    pub fn set_effect(&mut self, effect: Effect) -> Result<()> {
        effect.validate()?;
        self.effect = effect;
        Ok(())
    }

    /// Replace the effect in a consuming builder step.
    pub fn with_effect(mut self, effect: Effect) -> Result<Self> {
        self.set_effect(effect)?;
        Ok(self)
    }

    /// Return the start relationship.
    #[must_use]
    pub const fn start(&self) -> Start {
        self.start
    }

    /// Return the build duration.
    #[must_use]
    pub const fn duration(&self) -> Seconds {
        self.duration
    }

    /// Return the start delay.
    #[must_use]
    pub const fn delay(&self) -> Seconds {
        self.delay
    }

    /// Replace the duration with an already validated semantic value.
    pub const fn set_duration(&mut self, duration: Seconds) {
        self.duration = duration;
    }

    /// Replace the duration in a consuming builder step.
    #[must_use]
    pub const fn with_duration(mut self, duration: Seconds) -> Self {
        self.set_duration(duration);
        self
    }

    /// Replace the start relationship while preserving valid delay rules.
    pub fn set_start(&mut self, start: Start) -> Result<()> {
        if matches!(start, Start::OnClick | Start::WithPrevious) && self.delay != Seconds::ZERO {
            return Err(Error::InvalidBuildValue);
        }
        self.start = start;
        Ok(())
    }

    /// Replace the delay while preserving valid delay rules.
    pub fn set_delay(&mut self, delay: Seconds) -> Result<()> {
        if matches!(self.start, Start::OnClick | Start::WithPrevious) && delay != Seconds::ZERO {
            return Err(Error::InvalidBuildValue);
        }
        self.delay = delay;
        Ok(())
    }

    /// Set the start relationship in a consuming builder step.
    pub fn with_start(mut self, start: Start) -> Result<Self> {
        self.set_start(start)?;
        Ok(self)
    }

    /// Set the delay in a consuming builder step.
    pub fn with_delay(mut self, delay: Seconds) -> Result<Self> {
        self.set_delay(delay)?;
        Ok(self)
    }

    /// Validate the complete semantic combination before native publication.
    pub fn validate(&self) -> Result<()> {
        self.effect.validate()?;
        if self.duration == Seconds::ZERO {
            return Err(Error::InvalidBuildValue);
        }
        if matches!(self.start, Start::OnClick | Start::WithPrevious) && self.delay != Seconds::ZERO
        {
            return Err(Error::InvalidBuildValue);
        }
        Ok(())
    }
}

fn is_simple_effect_identifier(identifier: &str) -> bool {
    identifier.eq_ignore_ascii_case(APPEAR_EFFECT)
        || identifier.eq_ignore_ascii_case(DISSOLVE_EFFECT)
        || identifier.eq_ignore_ascii_case(MOVE_IN_EFFECT)
        || identifier.eq_ignore_ascii_case(SCALE_EFFECT)
        || identifier.eq_ignore_ascii_case(FADE_AND_SCALE_EFFECT)
}

fn is_action_identifier(identifier: &str) -> bool {
    matches!(
        identifier,
        ROTATE_ACTION_EFFECT | SCALE_ACTION_EFFECT | OPACITY_ACTION_EFFECT | MOTION_ACTION_EFFECT
    )
}

fn is_emphasis_identifier(identifier: &str) -> bool {
    matches!(
        identifier,
        BLINK_ACTION_EFFECT
            | BOUNCE_ACTION_EFFECT
            | FLIP_ACTION_EFFECT
            | JIGGLE_ACTION_EFFECT
            | POP_ACTION_EFFECT
            | PULSE_ACTION_EFFECT
    )
}

fn contains_ascii_case_insensitive(haystack: &str, needle: &[u8]) -> bool {
    haystack
        .as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn unknown_text_is_bounded_and_nul_free() {
        assert_eq!(UnknownText::new(""), Err(Error::EmptyIdentifier));
        assert_eq!(UnknownText::new("future\0effect"), Err(Error::NulString));
        assert_eq!(
            UnknownText::new("x".repeat(MAX_IDENTIFIER_BYTES + 1)),
            Err(Error::IdentifierTooLarge)
        );
        assert_eq!(
            UnknownText::new("future.effect").unwrap().as_str(),
            "future.effect"
        );
    }

    #[test]
    fn effects_preserve_unknown_identifiers_without_raw_fields() -> Result<()> {
        assert_eq!(Effect::from_identifier("APPEAR")?, Effect::Appear);
        let unknown = Effect::from_identifier("com.example.future")?;
        assert_eq!(unknown.identifier(), "com.example.future");
        assert_eq!(unknown, Effect::unknown("com.example.future")?);
        assert_eq!(Effect::unknown("appear"), Err(Error::NonCanonicalEffect));
        Ok(())
    }

    #[test]
    fn typed_effects_have_private_checked_parameters() -> Result<()> {
        let action = Action::Rotate(Rotation::new(
            810.0,
            RotationDirection::Clockwise,
            Acceleration::EaseIn,
        )?);
        assert_eq!(action.identifier(), ROTATE_ACTION_EFFECT);
        assert_eq!(Effect::action(action).identifier(), ROTATE_ACTION_EFFECT);
        assert_eq!(
            Rotation::new(f64::NAN, RotationDirection::Clockwise, Acceleration::None),
            Err(Error::InvalidBuildValue)
        );
        assert_eq!(
            Opacity::new(100.1, Acceleration::None),
            Err(Error::InvalidBuildValue)
        );
        assert_eq!(Emphasis::blink(0, true), Err(Error::InvalidBuildValue));
        assert_eq!(Emphasis::pulse(2, 0.0), Err(Error::InvalidBuildValue));
        Ok(())
    }

    #[test]
    fn paths_use_bounded_boxed_slices_and_reject_invalid_geometry() -> Result<()> {
        let path = Path::straight(320.0, -120.0)?;
        assert_eq!(path.subpaths().len(), 1);
        assert_eq!(path.subpaths()[0].nodes().len(), 2);
        assert_eq!(size_of::<Option<Box<[Node]>>>(), 2 * size_of::<usize>());
        assert_eq!(Path::straight(f64::NAN, 1.0), Err(Error::InvalidBuildValue));
        assert_eq!(
            Subpath::new(Vec::<Node>::new().into_boxed_slice(), false),
            Err(Error::InvalidBuildPath)
        );
        Ok(())
    }

    #[test]
    fn timing_curves_are_normalized_and_lossless() -> Result<()> {
        let curve = TimingCurve::linear();
        assert_eq!(curve.path().natural_width(), 1.0);
        assert_eq!(curve.path().subpaths()[0].nodes().len(), 2);
        let curve = TimingCurve::cubic(Point::new(0.2, 0.1)?, Point::new(0.8, 0.9)?)?;
        assert_eq!(curve.path().natural_height(), 1.0);
        assert_eq!(
            TimingCurve::cubic(Point::new(1.1, 0.0)?, Point::new(0.8, 0.9)?),
            Err(Error::InvalidBuildPath)
        );
        Ok(())
    }

    #[test]
    fn settings_enforce_start_delay_relationships() -> Result<()> {
        let effect = Effect::Appear;
        let duration = Seconds::new(0.5)?;
        let mut settings = Settings::new(effect, duration);
        assert_eq!(
            settings.clone().with_delay(Seconds::new(0.2)?),
            Err(Error::InvalidBuildValue)
        );
        settings.set_start(Start::AfterPrevious)?;
        settings.set_delay(Seconds::new(0.2)?)?;
        assert_eq!(settings.delay().as_f64(), 0.2);
        assert_eq!(settings.start(), Start::AfterPrevious);
        assert_eq!(
            settings.set_start(Start::OnClick),
            Err(Error::InvalidBuildValue)
        );
        assert_eq!(settings.start(), Start::AfterPrevious);
        assert_eq!(settings.set_delay(Seconds::new(0.3)?), Ok(()));
        assert_eq!(settings.delay().as_f64(), 0.3);
        assert_eq!(
            settings.set_start(Start::WithPrevious),
            Err(Error::InvalidBuildValue)
        );
        assert_eq!(settings.start(), Start::AfterPrevious);
        Ok(())
    }

    #[test]
    fn settings_reject_zero_duration_before_publication() {
        let settings = Settings::new(Effect::Appear, Seconds::ZERO);
        assert_eq!(settings.validate(), Err(Error::InvalidBuildValue));
    }
}
