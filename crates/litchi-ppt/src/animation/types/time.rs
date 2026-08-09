//! `PowerPoint` 2002 time-node and behavior data types.

use super::build::ChartBuildType;

/// Exact `PowerPoint` 2002 extended time-node container envelope.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct ExtendedTimeNode {
    /// Required first atom containing time-node attributes.
    pub atom: TimeNodeAtom,
    /// Optional typed properties immediately following the atom.
    pub properties: Option<TimeNodePropertyList>,
    /// The single animation behavior attached to a behavior node.
    pub behavior: Option<TimeNodeBehavior>,
    /// Media target attached to a media node.
    pub visual_target: Option<TimeVisualElement>,
    /// Optional repeated-subelement controls.
    pub iterate_data: Option<TimeIterateData>,
    /// Optional controls for a sequential node's children.
    pub sequence_data: Option<TimeSequenceData>,
    /// Begin or, for sequential nodes, next-child conditions.
    pub begin_conditions: Vec<TimeCondition>,
    /// End or, for sequential nodes, previous-child conditions.
    pub end_conditions: Vec<TimeCondition>,
    /// Optional child-stop synchronization condition.
    pub end_sync_condition: Option<TimeCondition>,
    /// Timing transformations applied to this node.
    pub modifiers: Vec<TimeModifier>,
    /// Subordinate effects whose starts depend on this node.
    pub sub_effects: Vec<TimeSubEffect>,
    /// Recursively nested time nodes.
    pub children: Vec<ExtendedTimeNode>,
}

/// The mutually exclusive behavior slots in an extended time node.
#[derive(Debug, Clone, PartialEq)]
pub enum TimeNodeBehavior {
    Animate(TimeAnimateBehavior),
    Color(TimeColorBehavior),
    Effect(TimeEffectBehavior),
    Motion(TimeMotionBehavior),
    Rotation(TimeRotationBehavior),
    Scale(TimeScaleBehavior),
    Set(TimeSetBehavior),
    Command(TimeCommandBehavior),
}

/// A subordinate behavior or media node attached to a master time node.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeSubEffect {
    pub atom: TimeNodeAtom,
    pub properties: Option<TimeNodePropertyList>,
    pub behavior: Option<TimeSubEffectBehavior>,
    pub visual_target: Option<TimeVisualElement>,
    pub begin_conditions: Vec<TimeCondition>,
    pub end_conditions: Vec<TimeCondition>,
    pub modifiers: Vec<TimeModifier>,
}

/// The mutually exclusive behavior slots supported by a subordinate effect.
#[derive(Debug, Clone, PartialEq)]
pub enum TimeSubEffectBehavior {
    Color(TimeColorBehavior),
    Set(TimeSetBehavior),
    Command(TimeCommandBehavior),
}

/// Exact fields controlled by a `PowerPoint` 2002 `TimeNodeAtom`.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TimeNodeAtom {
    /// Fill behavior when explicitly set; otherwise the format default applies.
    pub fill: Option<TimeNodeFill>,
    /// Restart behavior when explicitly set; otherwise the format default applies.
    pub restart: Option<TimeNodeRestart>,
    /// Node kind when explicitly set; otherwise `Parallel` applies.
    pub node_type: Option<TimeNodeKind>,
    /// Signed duration in milliseconds when explicitly set.
    pub duration_ms: Option<i32>,
}

/// Exact time-node kind values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeNodeKind {
    Parallel,
    Sequential,
    Behavior,
    Media,
}

impl TimeNodeKind {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::Parallel),
            1 => Some(Self::Sequential),
            3 => Some(Self::Behavior),
            4 => Some(Self::Media),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::Parallel => 0,
            Self::Sequential => 1,
            Self::Behavior => 3,
            Self::Media => 4,
        }
    }
}

/// Exact restart values, including the legacy alias for `Never`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeNodeRestart {
    Never,
    Always,
    WhenNotActive,
    NeverLegacy,
}

impl TimeNodeRestart {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::Never),
            1 => Some(Self::Always),
            2 => Some(Self::WhenNotActive),
            3 => Some(Self::NeverLegacy),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::Never => 0,
            Self::Always => 1,
            Self::WhenNotActive => 2,
            Self::NeverLegacy => 3,
        }
    }
}

/// Exact fill values, including the two legacy aliases.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeNodeFill {
    HoldUntilParentEnds,
    ResetWhenInactive,
    HoldUntilNext,
    HoldUntilParentEndsLegacy,
    ResetWhenInactiveLegacy,
}

/// Context in which a time-node property list occurs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimePropertyListContext {
    TimeNode,
    SubEffect,
}

/// Typed properties stored in `TimePropertyList4TimeNodeContainer`.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeNodePropertyList {
    pub properties: Vec<TimeNodeProperty>,
}

/// One time-node property, identified by its record instance.
#[derive(Debug, Clone, PartialEq)]
pub enum TimeNodeProperty {
    DisplayHidden(bool),
    MasterRelation(TimeMasterRelation),
    SubType,
    EffectId(i32),
    EffectDirection(i32),
    EffectType(TimeEffectType),
    AfterEffect(bool),
    SlideCount(i32),
    TimeFilter(String),
    EventFilter(String),
    HideWhenStopped(bool),
    GroupId(i32),
    EffectNodeType(TimeEffectNodeType),
    PlaceholderNode(bool),
    MediaVolume(f32),
    MediaMute(bool),
    ZoomToFullScreen(bool),
}

/// Relationship of a subordinate node to its master node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeMasterRelation {
    DoNotStart,
    StartWithMaster,
}

/// Animation effect category.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeEffectType {
    Entrance,
    Exit,
    Emphasis,
    MotionPath,
    ActionVerb,
    MediaCommand,
}

pub(crate) fn has_valid_time_effect_properties(properties: &[TimeNodeProperty]) -> bool {
    let mut maybe_effect_id = None;
    let mut maybe_direction = None;
    let mut maybe_effect_type = None;
    for property in properties {
        match property {
            TimeNodeProperty::EffectId(value) => maybe_effect_id = Some(*value),
            TimeNodeProperty::EffectDirection(value) => maybe_direction = Some(*value),
            TimeNodeProperty::EffectType(value) => maybe_effect_type = Some(*value),
            TimeNodeProperty::DisplayHidden(_)
            | TimeNodeProperty::MasterRelation(_)
            | TimeNodeProperty::SubType
            | TimeNodeProperty::AfterEffect(_)
            | TimeNodeProperty::SlideCount(_)
            | TimeNodeProperty::TimeFilter(_)
            | TimeNodeProperty::EventFilter(_)
            | TimeNodeProperty::HideWhenStopped(_)
            | TimeNodeProperty::GroupId(_)
            | TimeNodeProperty::EffectNodeType(_)
            | TimeNodeProperty::PlaceholderNode(_)
            | TimeNodeProperty::MediaVolume(_)
            | TimeNodeProperty::MediaMute(_)
            | TimeNodeProperty::ZoomToFullScreen(_) => {},
        }
    }
    let Some(effect_id) = maybe_effect_id else {
        return maybe_direction.is_none();
    };
    let Some(effect_type) = maybe_effect_type else {
        return false;
    };
    let valid_id = match effect_type {
        TimeEffectType::Entrance | TimeEffectType::Exit => (0..=0x3A).contains(&effect_id),
        TimeEffectType::Emphasis => (0..=0x24).contains(&effect_id),
        TimeEffectType::MotionPath => (0..=0x40).contains(&effect_id),
        TimeEffectType::MediaCommand => (0..=3).contains(&effect_id),
        TimeEffectType::ActionVerb => false,
    };
    valid_id
        && maybe_direction.is_none_or(|direction| {
            is_valid_time_effect_direction(effect_type, effect_id, direction)
        })
}

/// Role of a time node in the timing structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeEffectNodeType {
    ClickEffect,
    WithPrevious,
    AfterPrevious,
    MainSequence,
    InteractiveSequence,
    ClickParallel,
    WithGroup,
    AfterGroup,
    TimingRoot,
}

fn is_valid_time_effect_direction(
    effect_type: TimeEffectType,
    effect_id: i32,
    direction: i32,
) -> bool {
    match effect_type {
        TimeEffectType::Entrance | TimeEffectType::Exit => match effect_id {
            0x02 | 0x07 => matches!(direction, 1 | 2 | 3 | 4 | 6 | 8 | 9 | 12),
            0x03 | 0x05 | 0x0E | 0x13 => matches!(direction, 5 | 10),
            0x04 | 0x06 | 0x08 | 0x0D => matches!(direction, 0x10 | 0x20),
            0x0C | 0x16 => matches!(direction, 1 | 2 | 4 | 8),
            0x10 => matches!(direction, 0x15 | 0x1A | 0x25 | 0x2A),
            0x11 => matches!(direction, 1 | 2 | 4 | 8 | 10),
            0x12 => matches!(direction, 3 | 6 | 9 | 12),
            0x15 => matches!(direction, 1 | 2 | 3 | 4 | 8),
            0x17 => matches!(direction, 0x10 | 0x20 | 0x24 | 0x110 | 0x120 | 0x210),
            _ => true,
        },
        TimeEffectType::Emphasis => match effect_id {
            0x01 | 0x07 => matches!(direction, 1 | 2 | 6 | 10),
            // MS-PPT lists 0x01 for both instant and gradual for font color.
            0x03 => matches!(direction, 1 | 6 | 10),
            0x04 => matches!(direction, 1 | 2),
            0x05 => (0..=7).contains(&direction),
            _ => true,
        },
        TimeEffectType::MotionPath | TimeEffectType::MediaCommand => true,
        TimeEffectType::ActionVerb => false,
    }
}

/// Shared information used by all `PowerPoint` 2002 animation behaviors.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeBehavior {
    pub atom: TimeBehaviorAtom,
    /// Optional property names retained even when the atom marks them as ignored.
    pub attribute_names: Option<Vec<String>>,
    pub properties: Option<TimeBehaviorPropertyList>,
    pub target: TimeVisualElement,
}

/// Flags and composition mode from a `TimeBehaviorAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeBehaviorAtom {
    /// Explicit additive mode, or `None` when the file uses the default override mode.
    pub additive: Option<TimeBehaviorAdditive>,
    /// Whether the optional attribute-name list is meaningful.
    pub attribute_names_used: bool,
}

/// Composition mode for an animated value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeBehaviorAdditive {
    Override,
    Add,
}

/// A `PowerPoint` 2002 rotation behavior and its common target information.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeRotationBehavior {
    pub atom: TimeRotationBehaviorAtom,
    pub behavior: TimeBehavior,
}

/// Values controlled by a `TimeRotationBehaviorAtom`.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeRotationBehaviorAtom {
    pub by_degrees: Option<f32>,
    pub from_degrees: Option<f32>,
    pub to_degrees: Option<f32>,
    pub direction: Option<TimeRotationDirection>,
}

/// Direction used by a rotation behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeRotationDirection {
    Clockwise,
    CounterClockwise,
}

/// A `PowerPoint` 2002 scale behavior and its common target information.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeScaleBehavior {
    pub atom: TimeScaleBehaviorAtom,
    pub behavior: TimeBehavior,
}

/// Values controlled by a `TimeScaleBehaviorAtom`.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeScaleBehaviorAtom {
    pub by_percent: Option<(f32, f32)>,
    pub from_percent: Option<(f32, f32)>,
    pub to_percent: Option<(f32, f32)>,
    pub zoom_contents: Option<bool>,
}

/// A `PowerPoint` 2002 command behavior and its common target information.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeCommandBehavior {
    pub atom: TimeCommandBehaviorAtom,
    /// Optional command retained even when the atom marks it as ignored.
    pub command: Option<String>,
    pub behavior: TimeBehavior,
}

/// Flags and command kind from a `TimeCommandBehaviorAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeCommandBehaviorAtom {
    pub command_type: Option<TimeCommandBehaviorType>,
    pub command_used: bool,
}

/// Operation performed by a command behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeCommandBehaviorType {
    Event,
    Call,
    OleVerb,
}

/// Optional iteration controls for repeated sub-element effects.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeIterateData {
    pub interval: Option<u32>,
    pub iterate_type: Option<TimeIterateType>,
    pub direction: Option<TimeIterateDirection>,
    pub interval_type: Option<TimeIterateIntervalType>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeIterateType {
    AllAtOnce,
    ByWord,
    ByLetter,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeIterateDirection {
    Backward,
    Forward,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeIterateIntervalType {
    Milliseconds,
    TenthsOfAPercent,
}

/// Optional sequencing controls for a sequential time node.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeSequenceData {
    pub concurrent: Option<bool>,
    pub next_action: Option<TimeSequenceNextAction>,
    pub previous_action: Option<TimeSequencePreviousAction>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeSequenceNextAction {
    None,
    SeekToNaturalEnd,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeSequencePreviousAction {
    None,
    SkipTimedChildren,
}

/// A condition that controls activation or deactivation of a time node.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeCondition {
    pub condition_type: TimeConditionType,
    pub atom: TimeConditionAtom,
    /// Present if and only if `trigger_object` is `VisualElement`.
    pub visual_target: Option<TimeVisualElement>,
}

/// Fixed fields stored in a `TimeConditionAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeConditionAtom {
    pub trigger_object: TimeTriggerObject,
    pub trigger_event: TimeTriggerEvent,
    pub target_id: u32,
    pub delay_ms: i32,
}

/// Role of a condition within its containing time node.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeConditionType {
    None,
    Begin,
    End,
    Next,
    Previous,
    EndSync,
}

/// Kind of target participating in condition evaluation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeTriggerObject {
    None,
    VisualElement,
    TimeNode,
    RuntimeNodeReference,
}

/// Event that makes a time condition true.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeTriggerEvent {
    None,
    OnBegin,
    TimeNodeStart,
    TimeNodeEnd,
    MouseClick,
    MouseOver,
    OnNext,
    OnPrevious,
    StopAudio,
}

/// One `TimeModifierAtom`; values remain unsigned as required by MS-PPT.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TimeModifier {
    RepeatCount(u32),
    RepeatDuration(u32),
    Speed(u32),
    Accelerate(u32),
    Decelerate(u32),
    AutomaticReverse(u32),
}

/// Typed properties stored in `TimePropertyList4TimeBehavior`.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeBehaviorPropertyList {
    pub properties: Vec<TimeBehaviorProperty>,
}

/// One shared behavior property, identified by its record instance.
#[derive(Debug, Clone, PartialEq)]
pub enum TimeBehaviorProperty {
    UnknownPropertyList(String),
    RuntimeContext(String),
    MotionPathEditRelative(bool),
    ColorModel(TimeColorModel),
    ColorDirection(TimeColorDirection),
    Override,
    PathEditRotationAngle(f32),
    PathEditRotationX(f32),
    PathEditRotationY(f32),
    PointsTypes(String),
}

/// Color space used by a color behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeColorModel {
    Rgb,
    Hsl,
    Scheme,
}

/// Hue interpolation direction in HSL color space.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeColorDirection {
    Clockwise,
    CounterClockwise,
}

/// A `PowerPoint` 2002 color behavior and its common target information.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeColorBehavior {
    pub atom: TimeColorBehaviorAtom,
    pub behavior: TimeBehavior,
}

/// Values and property-use flags from a `TimeColorBehaviorAtom`.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeColorBehaviorAtom {
    pub by: Option<TimeAnimateColorBy>,
    pub from: Option<TimeAnimateColor>,
    pub to: Option<TimeAnimateColor>,
    pub color_space_used: bool,
    pub direction_used: bool,
}

/// Offset color used by a color animation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TimeAnimateColorBy {
    Rgb {
        red: i32,
        green: i32,
        blue: i32,
    },
    Hsl {
        hue: i32,
        saturation: i32,
        luminance: i32,
    },
    Scheme(u32),
}

/// Absolute color used by a color animation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TimeAnimateColor {
    Rgb { red: u32, green: u32, blue: u32 },
    Scheme(u32),
}

/// A `PowerPoint` 2002 image-transition behavior.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeEffectBehavior {
    pub atom: TimeEffectBehaviorAtom,
    /// Optional transition filter retained even when its use flag is clear.
    pub filter: Option<TimeEffectFilter>,
    /// Optional normalized progress retained even when its use flag is clear.
    pub progress: Option<f32>,
    /// Optional obsolete runtime context retained even when its use flag is clear.
    pub runtime_context: Option<String>,
    pub behavior: TimeBehavior,
}

/// Flags and transition direction from a `TimeEffectBehaviorAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeEffectBehaviorAtom {
    /// Explicit transition direction, or `None` for the default transition-in value.
    pub transition: Option<TimeEffectTransition>,
    pub filter_used: bool,
    pub progress_used: bool,
    pub runtime_context_used: bool,
}

/// Visibility direction of an image transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeEffectTransition {
    In,
    Out,
}

/// Transition filter supported by `TimeEffectBehaviorContainer`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeEffectFilter {
    BlindsHorizontal,
    BlindsVertical,
    BoxIn,
    BoxOut,
    CheckerboardAcross,
    CheckerboardDown,
    CircleIn,
    CircleOut,
    DiamondIn,
    DiamondOut,
    Dissolve,
    Fade,
    PlusIn,
    PlusOut,
    BarnInVertical,
    BarnInHorizontal,
    BarnOutVertical,
    BarnOutHorizontal,
    RandomBarHorizontal,
    RandomBarVertical,
    StripsDownLeft,
    StripsUpLeft,
    StripsDownRight,
    StripsUpRight,
    Wedge,
    Wheel1,
    Wheel2,
    Wheel3,
    Wheel4,
    Wheel8,
    WipeRight,
    WipeLeft,
    WipeUp,
    WipeDown,
}

impl TimeEffectFilter {
    pub(crate) fn parse(value: &str) -> Option<Self> {
        Some(match value {
            "blinds(horizontal)" => Self::BlindsHorizontal,
            "blinds(vertical)" => Self::BlindsVertical,
            "box(in)" => Self::BoxIn,
            "box(out)" => Self::BoxOut,
            "checkerboard(across)" => Self::CheckerboardAcross,
            "checkerboard(down)" => Self::CheckerboardDown,
            "circle(in)" => Self::CircleIn,
            "circle(out)" => Self::CircleOut,
            "diamond(in)" => Self::DiamondIn,
            "diamond(out)" => Self::DiamondOut,
            "dissolve" => Self::Dissolve,
            "fade" => Self::Fade,
            "plus(in)" => Self::PlusIn,
            "plus(out)" => Self::PlusOut,
            "barn(inVertical)" => Self::BarnInVertical,
            "barn(inHorizontal)" => Self::BarnInHorizontal,
            "barn(outVertical)" => Self::BarnOutVertical,
            "barn(outHorizontal)" => Self::BarnOutHorizontal,
            "randombar(horizontal)" => Self::RandomBarHorizontal,
            "randombar(vertical)" => Self::RandomBarVertical,
            "strips(downLeft)" => Self::StripsDownLeft,
            "strips(upLeft)" => Self::StripsUpLeft,
            "strips(downRight)" => Self::StripsDownRight,
            "strips(upRight)" => Self::StripsUpRight,
            "wedge" => Self::Wedge,
            "wheel(1)" => Self::Wheel1,
            "wheel(2)" => Self::Wheel2,
            "wheel(3)" => Self::Wheel3,
            "wheel(4)" => Self::Wheel4,
            "wheel(8)" => Self::Wheel8,
            "wipe(right)" => Self::WipeRight,
            "wipe(left)" => Self::WipeLeft,
            "wipe(up)" => Self::WipeUp,
            "wipe(down)" => Self::WipeDown,
            _ => return None,
        })
    }

    pub(crate) const fn as_str(self) -> &'static str {
        match self {
            Self::BlindsHorizontal => "blinds(horizontal)",
            Self::BlindsVertical => "blinds(vertical)",
            Self::BoxIn => "box(in)",
            Self::BoxOut => "box(out)",
            Self::CheckerboardAcross => "checkerboard(across)",
            Self::CheckerboardDown => "checkerboard(down)",
            Self::CircleIn => "circle(in)",
            Self::CircleOut => "circle(out)",
            Self::DiamondIn => "diamond(in)",
            Self::DiamondOut => "diamond(out)",
            Self::Dissolve => "dissolve",
            Self::Fade => "fade",
            Self::PlusIn => "plus(in)",
            Self::PlusOut => "plus(out)",
            Self::BarnInVertical => "barn(inVertical)",
            Self::BarnInHorizontal => "barn(inHorizontal)",
            Self::BarnOutVertical => "barn(outVertical)",
            Self::BarnOutHorizontal => "barn(outHorizontal)",
            Self::RandomBarHorizontal => "randombar(horizontal)",
            Self::RandomBarVertical => "randombar(vertical)",
            Self::StripsDownLeft => "strips(downLeft)",
            Self::StripsUpLeft => "strips(upLeft)",
            Self::StripsDownRight => "strips(downRight)",
            Self::StripsUpRight => "strips(upRight)",
            Self::Wedge => "wedge",
            Self::Wheel1 => "wheel(1)",
            Self::Wheel2 => "wheel(2)",
            Self::Wheel3 => "wheel(3)",
            Self::Wheel4 => "wheel(4)",
            Self::Wheel8 => "wheel(8)",
            Self::WipeRight => "wipe(right)",
            Self::WipeLeft => "wipe(left)",
            Self::WipeUp => "wipe(up)",
            Self::WipeDown => "wipe(down)",
        }
    }
}

/// A `PowerPoint` 2002 motion-path behavior.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeMotionBehavior {
    pub atom: TimeMotionBehaviorAtom,
    /// Optional path retained even when its use flag is clear.
    pub path: Option<String>,
    /// Optional obsolete integer record retained for round-trip fidelity.
    pub reserved: Option<i32>,
    pub behavior: TimeBehavior,
}

/// Values and property-use flags from a `TimeMotionBehaviorAtom`.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeMotionBehaviorAtom {
    pub by: Option<(f32, f32)>,
    pub from: Option<(f32, f32)>,
    pub to: Option<(f32, f32)>,
    pub origin: Option<TimeMotionOrigin>,
    pub path_used: bool,
    pub edit_rotation_used: bool,
    pub points_types_used: bool,
}

/// Wire-level origin of a motion path.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeMotionOrigin {
    Slide,
    SlideLegacy,
    ObjectCenter,
}

/// A `PowerPoint` 2002 behavior that assigns one property value.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeSetBehavior {
    pub atom: TimeSetBehaviorAtom,
    /// Optional value retained even when its use flag is clear.
    pub to: Option<String>,
    pub behavior: TimeBehavior,
}

/// Property-use flags and value type from a `TimeSetBehaviorAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TimeSetBehaviorAtom {
    pub to_used: bool,
    /// Explicit value type, or `None` for the default numeric type.
    pub value_type: Option<TimeAnimateValueType>,
}

/// Data type of a generic animated property value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeAnimateValueType {
    String,
    Number,
    Color,
}

/// A generic `PowerPoint` 2002 property animation.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeAnimateBehavior {
    pub atom: TimeAnimateBehaviorAtom,
    pub values: Option<TimeAnimationValueList>,
    /// Optional values retained even when their corresponding use flags are clear.
    pub by: Option<String>,
    pub from: Option<String>,
    pub to: Option<String>,
    pub behavior: TimeBehavior,
}

/// Calculation mode and use flags from a `TimeAnimateBehaviorAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "the bool fields mirror the independent `f*Used` flag bits of the fixed MS-PPT `TimeAnimateBehaviorAtom` layout, so they cannot be merged into enums without losing the bit-level mapping"
)]
pub struct TimeAnimateBehaviorAtom {
    /// Explicit interpolation mode, or `None` for linear interpolation.
    pub calculation_mode: Option<TimeAnimateCalculationMode>,
    pub by_used: bool,
    pub from_used: bool,
    pub to_used: bool,
    pub animation_values_used: bool,
    /// Explicit value type, or `None` for the default numeric type.
    pub value_type: Option<TimeAnimateValueType>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeAnimateCalculationMode {
    Discrete,
    Linear,
    Formula,
}

/// Key points in a generic property animation.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct TimeAnimationValueList {
    pub entries: Vec<TimeAnimationValue>,
}

/// One animation key point.
#[derive(Debug, Clone, PartialEq)]
pub struct TimeAnimationValue {
    /// Thousandths of the animation duration, or -1000 for even partitioning.
    pub time: i32,
    pub value: Option<TimeVariantValue>,
    pub formula: Option<String>,
}

/// Value kinds supported by a generic `TimeVariant` keyframe child.
#[derive(Debug, Clone, PartialEq)]
pub enum TimeVariantValue {
    Boolean(bool),
    Integer(i32),
    Float(f32),
    String(String),
}

/// Animation target stored in a `ClientVisualElementContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TimeVisualElement {
    Page,
    Sound {
        kind: TimeVisualElementKind,
        sound_id_ref: u32,
    },
    Shape {
        kind: TimeVisualElementKind,
        shape_id_ref: u32,
        data1: i32,
        data2: i32,
    },
    Chart {
        shape_id_ref: u32,
        build_type: ChartBuildType,
        element_index: i32,
    },
}

/// Portion of a visual object targeted by a behavior.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TimeVisualElementKind {
    Shape,
    Page,
    TextRange,
    Audio,
    Video,
    ChartElement,
    ShapeOnly,
    AllTextRange,
}

impl TimeVisualElementKind {
    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::Shape => 0,
            Self::Page => 1,
            Self::TextRange => 2,
            Self::Audio => 3,
            Self::Video => 4,
            Self::ChartElement => 5,
            Self::ShapeOnly => 6,
            Self::AllTextRange => 8,
        }
    }

    pub(crate) const fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::Shape),
            1 => Some(Self::Page),
            2 => Some(Self::TextRange),
            3 => Some(Self::Audio),
            4 => Some(Self::Video),
            5 => Some(Self::ChartElement),
            6 => Some(Self::ShapeOnly),
            8 => Some(Self::AllTextRange),
            _ => None,
        }
    }
}
