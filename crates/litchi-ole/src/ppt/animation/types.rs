//! Animation data types.
//!
//! Provides structures representing PowerPoint animation and build effects.

use super::motion_path::MotionPath;
use super::sound::AnimationSound;
use super::triggers::{InteractiveTrigger, IterationType, RepeatBehavior};
use crate::ppt::records::PptRecord;

/// Animation information for a slide or shape.
#[derive(Debug, Clone)]
pub struct AnimationInfo {
    /// PowerPoint 97 shape animation atom, when present.
    pub legacy_atom: Option<LegacyAnimationAtom>,
    /// Build list (order of appearance animations)
    pub build_list: Option<BuildInfo>,
    /// Time node containers for advanced animations
    pub time_nodes: Vec<TimeNodeContainer>,
    /// Sound associated with animation
    pub sound: Option<AnimationSound>,
    /// Interactive trigger
    pub trigger: Option<InteractiveTrigger>,
    /// Iteration type (for text animations)
    pub iteration: IterationType,
    /// Repeat behavior
    pub repeat: RepeatBehavior,
    /// After-effect color (for dim effects)
    pub after_effect_color: Option<u32>,
    /// Raw animation records for advanced parsing
    pub raw_records: Vec<PptRecord>,
}

impl Default for AnimationInfo {
    fn default() -> Self {
        Self::new()
    }
}

impl AnimationInfo {
    /// Create a new empty animation info.
    pub fn new() -> Self {
        Self {
            legacy_atom: None,
            build_list: None,
            time_nodes: Vec::new(),
            sound: None,
            trigger: None,
            iteration: IterationType::default(),
            repeat: RepeatBehavior::default(),
            after_effect_color: None,
            raw_records: Vec::new(),
        }
    }

    /// Check if this slide has any animations.
    pub fn has_animations(&self) -> bool {
        self.legacy_atom
            .as_ref()
            .is_some_and(|atom| atom.build_type != LegacyAnimationBuild::NoBuild)
            || self.build_list.is_some()
            || !self.time_nodes.is_empty()
    }

    /// Get the number of animated objects.
    pub fn animation_count(&self) -> usize {
        let legacy_count = usize::from(
            self.legacy_atom
                .as_ref()
                .is_some_and(|atom| atom.build_type != LegacyAnimationBuild::NoBuild),
        );
        let build_count = self
            .build_list
            .as_ref()
            .map(|b| b.builds.len())
            .unwrap_or(0);
        legacy_count + build_count + self.time_nodes.len()
    }
}

/// Animation metadata associated with a legacy PowerPoint shape.
#[derive(Debug, Clone)]
pub struct ShapeAnimation {
    /// OfficeArt shape identifier.
    pub shape_id: u32,
    /// Parsed, inert animation metadata.
    pub animation: AnimationInfo,
}

/// PowerPoint 97 paragraph/chart build behavior stored in `AnimationInfoAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegacyAnimationBuild {
    FollowMaster,
    #[default]
    NoBuild,
    OneBuild,
    Level1,
    Level2,
    Level3,
    Level4,
    Level5,
    GraphBySeries,
    GraphByCategory,
    GraphByElementInSeries,
    GraphByElementInCategory,
}

impl LegacyAnimationBuild {
    pub(crate) fn parse(value: u8) -> Option<Self> {
        match value {
            0xFE => Some(Self::FollowMaster),
            0x00 => Some(Self::NoBuild),
            0x01 => Some(Self::OneBuild),
            0x02 => Some(Self::Level1),
            0x03 => Some(Self::Level2),
            0x04 => Some(Self::Level3),
            0x05 => Some(Self::Level4),
            0x06 => Some(Self::Level5),
            0x07 => Some(Self::GraphBySeries),
            0x08 => Some(Self::GraphByCategory),
            0x09 => Some(Self::GraphByElementInSeries),
            0x0A => Some(Self::GraphByElementInCategory),
            _ => None,
        }
    }

    pub(crate) const fn as_u8(self) -> u8 {
        match self {
            Self::FollowMaster => 0xFE,
            Self::NoBuild => 0x00,
            Self::OneBuild => 0x01,
            Self::Level1 => 0x02,
            Self::Level2 => 0x03,
            Self::Level3 => 0x04,
            Self::Level4 => 0x05,
            Self::Level5 => 0x06,
            Self::GraphBySeries => 0x07,
            Self::GraphByCategory => 0x08,
            Self::GraphByElementInSeries => 0x09,
            Self::GraphByElementInCategory => 0x0A,
        }
    }
}

/// PowerPoint 97 animation effect code.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegacyAnimationEffect {
    #[default]
    Cut,
    Random,
    Blinds,
    Checker,
    Cover,
    Dissolve,
    Fade,
    Pull,
    RandomBars,
    Strips,
    Wipe,
    Zoom,
    Fly,
    Split,
    Flash,
    Diamond,
    Plus,
    Wedge,
    Wheel,
    Circle,
}

impl LegacyAnimationEffect {
    pub(crate) fn parse(value: u8) -> Option<Self> {
        match value {
            0x00 => Some(Self::Cut),
            0x01 => Some(Self::Random),
            0x02 => Some(Self::Blinds),
            0x03 => Some(Self::Checker),
            0x04 => Some(Self::Cover),
            0x05 => Some(Self::Dissolve),
            0x06 => Some(Self::Fade),
            0x07 => Some(Self::Pull),
            0x08 => Some(Self::RandomBars),
            0x09 => Some(Self::Strips),
            0x0A => Some(Self::Wipe),
            0x0B => Some(Self::Zoom),
            0x0C => Some(Self::Fly),
            0x0D => Some(Self::Split),
            0x0E => Some(Self::Flash),
            0x11 => Some(Self::Diamond),
            0x12 => Some(Self::Plus),
            0x13 => Some(Self::Wedge),
            0x1A => Some(Self::Wheel),
            0x1B => Some(Self::Circle),
            _ => None,
        }
    }

    pub(crate) const fn as_u8(self) -> u8 {
        match self {
            Self::Cut => 0x00,
            Self::Random => 0x01,
            Self::Blinds => 0x02,
            Self::Checker => 0x03,
            Self::Cover => 0x04,
            Self::Dissolve => 0x05,
            Self::Fade => 0x06,
            Self::Pull => 0x07,
            Self::RandomBars => 0x08,
            Self::Strips => 0x09,
            Self::Wipe => 0x0A,
            Self::Zoom => 0x0B,
            Self::Fly => 0x0C,
            Self::Split => 0x0D,
            Self::Flash => 0x0E,
            Self::Diamond => 0x11,
            Self::Plus => 0x12,
            Self::Wedge => 0x13,
            Self::Wheel => 0x1A,
            Self::Circle => 0x1B,
        }
    }

    pub(crate) const fn accepts_direction(self, direction: u8) -> bool {
        match self {
            Self::Cut => direction <= 2,
            Self::Random => true,
            Self::Blinds | Self::Checker | Self::RandomBars | Self::Zoom => direction <= 1,
            Self::Cover | Self::Pull => direction <= 7,
            Self::Dissolve
            | Self::Fade
            | Self::Diamond
            | Self::Plus
            | Self::Wedge
            | Self::Circle => direction == 0,
            Self::Strips => direction >= 4 && direction <= 7,
            Self::Wipe | Self::Split => direction <= 3,
            Self::Fly => direction <= 0x1C,
            Self::Flash => direction <= 2,
            Self::Wheel => matches!(direction, 1 | 2 | 3 | 4 | 8),
        }
    }
}

/// Text subdivision behavior in a PowerPoint 97 animation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum LegacyTextBuildSubEffect {
    #[default]
    AllAtOnce,
    ByWord,
    ByCharacter,
}

impl LegacyTextBuildSubEffect {
    pub(crate) fn parse(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::AllAtOnce),
            1 => Some(Self::ByWord),
            2 => Some(Self::ByCharacter),
            _ => None,
        }
    }

    pub(crate) const fn as_u8(self) -> u8 {
        match self {
            Self::AllAtOnce => 0,
            Self::ByWord => 1,
            Self::ByCharacter => 2,
        }
    }
}

/// Exact 28-byte payload of an `[MS-PPT]` `AnimationInfoAtom`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct LegacyAnimationAtom {
    pub dim_color: u32,
    pub reverse: bool,
    pub automatic: bool,
    pub has_sound: bool,
    pub stop_sound: bool,
    pub play: bool,
    pub synchronous: bool,
    pub hide_while_not_playing: bool,
    pub animate_background: bool,
    pub sound_id_ref: u32,
    pub delay_time_ms: i32,
    pub order_id: i16,
    pub slide_count: u16,
    pub build_type: LegacyAnimationBuild,
    pub effect: LegacyAnimationEffect,
    pub effect_direction: u8,
    pub after_effect: AfterEffect,
    pub text_build_sub_effect: LegacyTextBuildSubEffect,
    pub ole_verb: u8,
}

impl Default for LegacyAnimationAtom {
    fn default() -> Self {
        Self {
            dim_color: 0,
            reverse: false,
            automatic: false,
            has_sound: false,
            stop_sound: false,
            play: false,
            synchronous: false,
            hide_while_not_playing: false,
            animate_background: false,
            sound_id_ref: 0,
            delay_time_ms: 0,
            order_id: 0,
            slide_count: 1,
            build_type: LegacyAnimationBuild::NoBuild,
            effect: LegacyAnimationEffect::Cut,
            effect_direction: 0,
            after_effect: AfterEffect::None,
            text_build_sub_effect: LegacyTextBuildSubEffect::AllAtOnce,
            ole_verb: 0,
        }
    }
}

/// High-level build list information used by animation authoring APIs.
#[derive(Debug, Clone, PartialEq)]
pub struct BuildInfo {
    /// Individual build items.
    pub builds: Vec<BuildLevel>,
}

impl BuildInfo {
    /// Create a new empty build info.
    pub fn new() -> Self {
        Self { builds: Vec::new() }
    }

    /// Add a build item.
    pub fn add_build(&mut self, build: BuildLevel) {
        self.builds.push(build);
    }
}

/// Exact PowerPoint 2002 build-list record for a slide.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct BuildList {
    /// Paragraph, chart, and diagram build subcontainers in file order.
    pub builds: Vec<BuildListEntry>,
}

/// PowerPoint 2002 animation metadata stored in a slide's `___PPT10` tag.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct SlideAnimationExtension {
    /// Optional root animation timing tree.
    pub time_node: Option<ExtendedTimeNode>,
    /// Optional shape build list.
    pub build_list: Option<BuildList>,
}

impl BuildList {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn add_build(&mut self, build: BuildListEntry) {
        self.builds.push(build);
    }
}

impl Default for BuildInfo {
    fn default() -> Self {
        Self::new()
    }
}

/// Shared build kind stored in a PowerPoint 2002 `BuildAtom`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum BuildKind {
    Paragraph,
    Chart,
    Diagram,
}

impl BuildKind {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            1 => Some(Self::Paragraph),
            2 => Some(Self::Chart),
            3 => Some(Self::Diagram),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::Paragraph => 1,
            Self::Chart => 2,
            Self::Diagram => 3,
        }
    }
}

/// Exact fields shared by paragraph, chart, and diagram build containers.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct BuildAtom {
    pub build_id: u32,
    pub shape_id_ref: u32,
    pub expanded: bool,
    pub ui_expanded: bool,
}

/// Paragraph build mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ParagraphBuildType {
    AllAtOnce,
    BuildByNthLevel,
    CustomBuild,
    AsAWhole,
}

impl ParagraphBuildType {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::AllAtOnce),
            1 => Some(Self::BuildByNthLevel),
            2 => Some(Self::CustomBuild),
            3 => Some(Self::AsAWhole),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::AllAtOnce => 0,
            Self::BuildByNthLevel => 1,
            Self::CustomBuild => 2,
            Self::AsAWhole => 3,
        }
    }
}

/// Exact paragraph-specific build atom fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphBuildAtom {
    pub build_type: ParagraphBuildType,
    pub build_level: u32,
    pub animate_background: bool,
    pub reverse: bool,
    pub user_set_animate_background: bool,
    pub automatic: bool,
    pub delay_time_ms: u32,
}

/// Template effect for one paragraph level.
#[derive(Debug, Clone, PartialEq)]
pub struct ParagraphBuildLevel {
    /// Paragraph level, at most 9.
    pub level: u32,
    /// Inert extended time node retained without executing actions.
    pub time_node: ExtendedTimeNode,
}

/// Text paragraph build subcontainer.
#[derive(Debug, Clone, PartialEq)]
pub struct ParagraphBuild {
    pub atom: BuildAtom,
    pub paragraph: ParagraphBuildAtom,
    pub levels: Vec<ParagraphBuildLevel>,
}

/// Chart build mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartBuildType {
    AsOneObject,
    BySeries,
    ByCategory,
    ByElementInSeries,
    ByElementInCategory,
    Custom,
}

impl ChartBuildType {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::AsOneObject),
            1 => Some(Self::BySeries),
            2 => Some(Self::ByCategory),
            3 => Some(Self::ByElementInSeries),
            4 => Some(Self::ByElementInCategory),
            5 => Some(Self::Custom),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::AsOneObject => 0,
            Self::BySeries => 1,
            Self::ByCategory => 2,
            Self::ByElementInSeries => 3,
            Self::ByElementInCategory => 4,
            Self::Custom => 5,
        }
    }
}

/// Chart-specific build atom fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartBuildAtom {
    pub build_type: ChartBuildType,
    pub animate_background: bool,
}

/// Chart build subcontainer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartBuild {
    pub atom: BuildAtom,
    pub chart: ChartBuildAtom,
}

/// Diagram build mode.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DiagramBuildType {
    AsOneObject,
    DepthByNode,
    DepthByBranch,
    BreadthByNode,
    BreadthByLevel,
    Clockwise,
    ClockwiseIn,
    ClockwiseOut,
    CounterClockwise,
    CounterClockwiseIn,
    CounterClockwiseOut,
    InByRing,
    OutByRing,
    Up,
    Down,
    AllAtOnce,
    Custom,
}

impl DiagramBuildType {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0x00 => Some(Self::AsOneObject),
            0x01 => Some(Self::DepthByNode),
            0x02 => Some(Self::DepthByBranch),
            0x03 => Some(Self::BreadthByNode),
            0x04 => Some(Self::BreadthByLevel),
            0x05 => Some(Self::Clockwise),
            0x06 => Some(Self::ClockwiseIn),
            0x07 => Some(Self::ClockwiseOut),
            0x08 => Some(Self::CounterClockwise),
            0x09 => Some(Self::CounterClockwiseIn),
            0x0A => Some(Self::CounterClockwiseOut),
            0x0B => Some(Self::InByRing),
            0x0C => Some(Self::OutByRing),
            0x0D => Some(Self::Up),
            0x0E => Some(Self::Down),
            0x0F => Some(Self::AllAtOnce),
            0x10 => Some(Self::Custom),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::AsOneObject => 0x00,
            Self::DepthByNode => 0x01,
            Self::DepthByBranch => 0x02,
            Self::BreadthByNode => 0x03,
            Self::BreadthByLevel => 0x04,
            Self::Clockwise => 0x05,
            Self::ClockwiseIn => 0x06,
            Self::ClockwiseOut => 0x07,
            Self::CounterClockwise => 0x08,
            Self::CounterClockwiseIn => 0x09,
            Self::CounterClockwiseOut => 0x0A,
            Self::InByRing => 0x0B,
            Self::OutByRing => 0x0C,
            Self::Up => 0x0D,
            Self::Down => 0x0E,
            Self::AllAtOnce => 0x0F,
            Self::Custom => 0x10,
        }
    }
}

/// Diagram-specific build atom fields.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagramBuildAtom {
    pub build_type: DiagramBuildType,
}

/// Diagram build subcontainer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DiagramBuild {
    pub atom: BuildAtom,
    pub diagram: DiagramBuildAtom,
}

/// One spec-defined PowerPoint 2002 build-list child.
#[derive(Debug, Clone, PartialEq)]
pub enum BuildListEntry {
    Paragraph(ParagraphBuild),
    Chart(ChartBuild),
    Diagram(DiagramBuild),
}

/// Exact PowerPoint 2002 extended time-node container envelope.
#[derive(Debug, Clone, PartialEq)]
pub struct ExtendedTimeNode {
    /// Required first atom containing time-node attributes.
    pub atom: TimeNodeAtom,
    /// Optional typed properties immediately following the atom.
    pub properties: Option<TimeNodePropertyList>,
    /// Remaining property, behavior, condition, modifier, and child records.
    pub children: Vec<PptRecord>,
}

/// Exact fields controlled by a PowerPoint 2002 `TimeNodeAtom`.
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

/// Shared information used by all PowerPoint 2002 animation behaviors.
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

pub(crate) fn is_valid_runtime_context(value: &str) -> bool {
    fn valid_version(value: &str) -> bool {
        let mut components = value.split('.');
        let first = components.next().unwrap_or_default();
        !first.is_empty()
            && first.bytes().all(|byte| byte.is_ascii_digit())
            && components.next().is_none_or(|second| {
                !second.is_empty()
                    && second.bytes().all(|byte| byte.is_ascii_digit())
                    && components.next().is_none()
            })
    }

    fn valid_atom(atom: &str) -> bool {
        fn valid_relation(value: &str) -> bool {
            value == "!"
                || ["gte", "gt", "lte", "lt"]
                    .iter()
                    .any(|relation| value.eq_ignore_ascii_case(relation))
        }

        if atom.is_empty()
            || atom.starts_with(' ')
            || atom.ends_with(' ')
            || atom
                .chars()
                .any(|character| character.is_whitespace() && character != ' ')
        {
            return false;
        }
        let fields = atom.split_ascii_whitespace().collect::<Vec<_>>();
        match fields.as_slice() {
            [app] => app.eq_ignore_ascii_case("ppt"),
            [first, second] => {
                (first.eq_ignore_ascii_case("ppt") && valid_version(second))
                    || (valid_relation(first) && second.eq_ignore_ascii_case("ppt"))
            },
            [relation, app, version] => {
                valid_relation(relation)
                    && app.eq_ignore_ascii_case("ppt")
                    && valid_version(version)
            },
            _ => false,
        }
    }

    let sequence = value.strip_suffix(';').unwrap_or(value);
    !sequence.is_empty() && sequence.split(';').all(valid_atom)
}

pub(crate) fn is_valid_time_points_types(value: &str) -> bool {
    value
        .bytes()
        .all(|byte| matches!(byte, b'A' | b'a' | b'F' | b'f' | b'T' | b't' | b'S' | b's'))
}

pub(crate) fn is_valid_time_filter(value: &str) -> bool {
    fn normalized_time(value: &str) -> bool {
        value == "1.0"
            || value.strip_prefix("0.").is_some_and(|digits| {
                !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit())
            })
    }

    !value.is_empty()
        && value.split(';').all(|entry| {
            let mut fields = entry.split(',');
            matches!(
                (fields.next(), fields.next(), fields.next()),
                (Some(time), Some(transformed), None)
                    if normalized_time(time) && normalized_time(transformed)
            )
        })
}

impl TimeNodeFill {
    pub(crate) fn parse(value: u32) -> Option<Self> {
        match value {
            0 => Some(Self::HoldUntilParentEnds),
            1 => Some(Self::ResetWhenInactive),
            2 => Some(Self::HoldUntilNext),
            3 => Some(Self::HoldUntilParentEndsLegacy),
            4 => Some(Self::ResetWhenInactiveLegacy),
            _ => None,
        }
    }

    pub(crate) const fn as_u32(self) -> u32 {
        match self {
            Self::HoldUntilParentEnds => 0,
            Self::ResetWhenInactive => 1,
            Self::HoldUntilNext => 2,
            Self::HoldUntilParentEndsLegacy => 3,
            Self::ResetWhenInactiveLegacy => 4,
        }
    }
}

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
    /// Fade out (same as FadeIn but exit type)
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
    pub raw_record: Option<PptRecord>,
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
